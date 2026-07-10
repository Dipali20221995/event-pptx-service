const express = require('express');
const multer = require('multer');
const PptxGenJS = require('pptxgenjs');
const sharp = require('sharp');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

// Simple shared-secret auth so randoms on the internet can't hit your endpoint.
// Set API_KEY as an environment variable on your host; n8n sends it as a header.
const API_KEY = process.env.API_KEY || '';

app.get('/', (req, res) => res.send('Event PPTX service is running.'));

// n8n's "Compress Image (External)" node sends the raw image bytes directly
// as the POST body (contentType: binaryData), not as multipart/form-data,
// so this route reads the raw body rather than using multer.
app.post('/compress-image', express.raw({ type: '*/*', limit: '25mb' }), async (req, res) => {
  try {
    if (API_KEY) {
      const provided = req.header('x-api-key') || '';
      if (provided !== API_KEY) {
        return res.status(401).json({ error: 'Unauthorized' });
      }
    }

    if (!req.body || !req.body.length) {
      return res.status(400).json({ error: 'Missing image body' });
    }

    const quality = parseInt(req.query.quality, 10) || 80;
    const maxWidth = parseInt(req.query.maxWidth, 10) || 1280;
    const maxHeight = parseInt(req.query.maxHeight, 10) || 1280;

    const buf = await sharp(req.body)
      .resize(maxWidth, maxHeight, { fit: 'inside', withoutEnlargement: true })
      .jpeg({ quality })
      .toBuffer();

    res.set('Content-Type', 'image/jpeg');
    res.send(buf);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

app.post('/generate-pptx', upload.any(), async (req, res) => {
  try {
    if (API_KEY) {
      const provided = req.header('x-api-key') || '';
      if (provided !== API_KEY) {
        return res.status(401).json({ error: 'Unauthorized' });
      }
    }

    if (!req.body.payload) {
      return res.status(400).json({ error: 'Missing "payload" field (JSON string)' });
    }

    const inputData = JSON.parse(req.body.payload);
    const slides = inputData.slides || [];
    const eventData = inputData.eventData || {};

    // ---- classify uploaded files into logo vs reference images ----
    const files = req.files || [];
    let logoFile = null;
    const refFiles = [];
    for (const f of files) {
      const nameHint = ((f.fieldname || '') + (f.originalname || '')).toLowerCase();
      if (nameHint.includes('logo') && !logoFile) {
        logoFile = f;
      } else {
        refFiles.push(f);
      }
    }

    const buf = await buildPresentation({ inputData, slides, eventData, logoFile, refFiles });

    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.set('Content-Disposition', `attachment; filename="presentation.pptx"`);
    res.send(buf);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

async function buildPresentation({ inputData, slides, eventData, logoFile, refFiles }) {
  // ── THEME COLOR SYSTEM ─────────────────────────────────────────
  const accent    = (inputData.accentColor    || '#B38E58').replace('#', '');
  const secondary = (inputData.secondaryColor || '#D4B896').replace('#', '');
  const bgColor   = (inputData.bgColor        || '#FBF7F2').replace('#', '');
  const darkColor = (inputData.darkColor      || '#3D2820').replace('#', '');
  const white     = 'FFFFFF';
  const textDark  = '2C2C2C';
  const textGrey  = '666666';

  // ── FIXED DESIGN CONSTANTS ─────────────────────────────────────
  const FONT   = 'Calibri';
  const HDR_H  = 0.75;
  const FTR_Y  = 5.38;
  const FTR_H  = 0.25;
  const LOGO_X = 8.78;
  const LOGO_Y = 0.09;
  const LOGO_W = 1.05;
  const LOGO_H = 0.54;
  const ML     = 0.35;
  const CTY    = HDR_H + 0.15;

  // ── PROCESS IMAGES ────────────────────────────────────────────
  const refImgs = [];
  let logoB64 = null;

  if (logoFile) {
    try {
      const lb = await sharp(logoFile.buffer)
        .resize(200, 100, { fit: 'inside', background: { r: 255, g: 255, b: 255, alpha: 0 } })
        .png()
        .toBuffer();
      logoB64 = 'data:image/png;base64,' + lb.toString('base64');
    } catch (e) { /* skip bad logo */ }
  }

  for (const f of refFiles) {
    try {
      // 'attention' crop strategy analyzes the image (edges, skin tones,
      // saturation) to keep the most visually interesting region in frame,
      // instead of blindly cutting to the exact center — this noticeably
      // reduces awkward crops on portrait or off-center photos. Resolution
      // and quality bumped up for a crisper result at typical slide sizes.
      const rb = await sharp(f.buffer)
        .resize(1600, 900, { fit: 'cover', position: 'attention' })
        .modulate({ brightness: 1.1, saturation: 0.85 })
        .jpeg({ quality: 92 })
        .toBuffer();
      refImgs.push('data:image/jpeg;base64,' + rb.toString('base64'));
    } catch (e) { /* skip bad image */ }
  }

  // Per-slide targeted image selection. Each slide object (sd) may carry an
  // `imageIndex` chosen upstream (by the AI planning step) that refers to the
  // position of a specific uploaded reference image. If it's missing, out of
  // range, or null, the slide simply gets no image (existing placeholder /
  // no-image branches below handle that gracefully) instead of forcing
  // whatever image happens to be "next" in a round-robin.
  const pickImg = (sd) => {
    const idx = sd && sd.imageIndex;
    if (typeof idx === 'number' && idx >= 0 && idx < refImgs.length) {
      return refImgs[idx];
    }
    return null;
  };

  // ── PPTX ──────────────────────────────────────────────────────
  // Document metadata fields (title/company/etc.) are written into the
  // file's internal XML without pptxgenjs escaping them automatically the
  // way it escapes text added to slides. A raw '&', '<', '>', or '"' in a
  // client/event name breaks that XML and makes PowerPoint report the file
  // as corrupt. Escaping here fixes that for ANY such character, not just
  // the one that happened to show up so far.
  const escapeXml = (s) => String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = escapeXml(eventData.eventName || 'Event Presentation');
  pptx.company = escapeXml(eventData.clientName || '');

  const addLogo = (slide) => {
    if (logoB64) {
      slide.addImage({ data: logoB64, x: LOGO_X, y: LOGO_Y, w: LOGO_W, h: LOGO_H });
    } else {
      slide.addShape(pptx.ShapeType.rect, { x: LOGO_X, y: LOGO_Y, w: LOGO_W, h: LOGO_H, fill: { color: white, alpha: 80 }, line: { color: white, width: 0.4 } });
      slide.addText('LOGO', { x: LOGO_X, y: LOGO_Y, w: LOGO_W, h: LOGO_H, fontSize: 9, color: accent, align: 'center', fontFace: FONT, bold: true, valign: 'middle' });
    }
  };

  const addFooter = (slide, num) => {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: FTR_Y, w: 10, h: FTR_H, fill: { color: accent } });
    slide.addText(eventData.clientName || '', { x: ML, y: FTR_Y + 0.04, w: 4, h: 0.18, fontSize: 8, color: white, fontFace: FONT });
    slide.addText(eventData.eventName || '', { x: 3.5, y: FTR_Y + 0.04, w: 3.5, h: 0.18, fontSize: 8, color: white, fontFace: FONT, align: 'center' });
    if (num) slide.addText(String(num), { x: 9.3, y: FTR_Y + 0.04, w: 0.55, h: 0.18, fontSize: 8, color: white, align: 'right', fontFace: FONT });
  };

  const addHeader = (slide, title) => {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: HDR_H, fill: { color: accent } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.08, h: HDR_H, fill: { color: secondary } });
    slide.addText((title || '').toUpperCase(), { x: 0.22, y: 0.1, w: LOGO_X - 0.3, h: HDR_H - 0.2, fontSize: 22, bold: true, color: white, fontFace: FONT, charSpacing: 1, valign: 'middle' });
    addLogo(slide);
  };

  const addPH = (slide, x, y, w, h, label) => {
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: 'EDEAE3' }, line: { color: accent, width: 1.2, dashType: 'dash' } });
    slide.addText('\uD83D\uDCF8', { x, y: y + h / 2 - 0.3, w, h: 0.5, fontSize: 20, align: 'center' });
    slide.addText('[ ' + (label || 'Add Image Here') + ' ]', { x, y: y + h / 2 + 0.25, w, h: 0.32, fontSize: 9.5, color: accent, align: 'center', fontFace: FONT, italic: true });
  };

  const accentLine = (slide, x, y, w) => slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.04, fill: { color: accent } });
  const secLine = (slide, x, y, w) => slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.04, fill: { color: secondary } });

  const buildCover = (sd) => {
    const slide = pptx.addSlide();
    const img = pickImg(sd);
    if (img) {
      slide.addImage({ data: img, x: 4.8, y: 0, w: 5.2, h: 5.63 });
      slide.addShape(pptx.ShapeType.rect, { x: 4.8, y: 0, w: 5.2, h: 5.63, fill: { color: darkColor, alpha: 15 } });
    } else {
      slide.background = { fill: bgColor };
      addPH(slide, 4.9, 0.15, 5.0, 5.3, 'Cover Image');
    }
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 4.75, h: 5.63, fill: { color: darkColor } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 4.75, h: 0.12, fill: { color: secondary } });
    slide.addShape(pptx.ShapeType.rect, { x: 4.62, y: 0, w: 0.13, h: 5.63, fill: { color: accent } });
    if (logoB64) slide.addImage({ data: logoB64, x: 0.4, y: 0.22, w: 1.5, h: 0.72 });
    else slide.addText('LOGO', { x: 0.4, y: 0.22, w: 1.5, h: 0.5, fontSize: 10, color: accent, fontFace: FONT, bold: true });
    slide.addText((eventData.eventName || '').toUpperCase(), { x: 0.42, y: 1.1, w: 4.1, h: 1.7, fontSize: 32, bold: true, color: white, fontFace: FONT, charSpacing: 1, wrap: true });
    accentLine(slide, 0.42, 2.95, 3.5);
    secLine(slide, 0.42, 3.02, 2.0);
    slide.addText('Presented to', { x: 0.42, y: 3.15, w: 4.1, h: 0.3, fontSize: 10, color: 'AAAAAA', fontFace: FONT });
    slide.addText(eventData.clientName || '', { x: 0.42, y: 3.45, w: 4.1, h: 0.42, fontSize: 15, bold: true, color: white, fontFace: FONT });
    slide.addText('\uD83D\uDCC5 ' + (eventData.eventDate || ''), { x: 0.42, y: 3.98, w: 4.1, h: 0.35, fontSize: 11, color: 'BBBBBB', fontFace: FONT });
    slide.addText('\uD83D\uDCCD ' + (eventData.venue || ''), { x: 0.42, y: 4.35, w: 4.1, h: 0.35, fontSize: 11, color: 'BBBBBB', fontFace: FONT });
    if (eventData.theme) {
      slide.addShape(pptx.ShapeType.rect, { x: 0.42, y: 4.78, w: 3.2, h: 0.4, fill: { color: accent } });
      slide.addText('\u2726 ' + (eventData.theme || ''), { x: 0.44, y: 4.8, w: 3.16, h: 0.36, fontSize: 11, color: white, fontFace: FONT, bold: true });
    }
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.45, w: 4.75, h: 0.18, fill: { color: accent } });
  };

  const buildContent = (sd, useImg) => {
    const slide = pptx.addSlide();
    slide.background = { fill: bgColor };
    addHeader(slide, sd.title);
    const img = useImg ? pickImg(sd) : null;
    const TW = 5.65;
    const IX = 6.08, IY = CTY + 0.05, IW = 3.65, IH = 4.35;

    if (sd.subtitle) slide.addText(sd.subtitle, { x: ML, y: CTY, w: TW, h: 0.38, fontSize: 13, color: accent, italic: true, fontFace: FONT });
    accentLine(slide, ML, CTY + 0.42, 2.4);

    const BY = sd.bodyText ? CTY + 0.55 : CTY + 0.48;
    if (sd.bodyText) slide.addText(sd.bodyText, { x: ML, y: CTY + 0.52, w: TW, h: 0.65, fontSize: 11.5, color: textGrey, fontFace: FONT, wrap: true });

    (sd.bulletPoints || []).forEach((pt, i) => {
      const y = BY + (sd.bodyText ? 0.72 : 0) + i * 0.55;
      slide.addShape(pptx.ShapeType.rect, { x: ML, y: y + 0.1, w: 0.06, h: 0.3, fill: { color: accent } });
      slide.addShape(pptx.ShapeType.rect, { x: ML + 0.07, y: y + 0.22, w: TW - 0.1, h: 0.02, fill: { color: secondary, alpha: 80 } });
      slide.addText(pt, { x: ML + 0.18, y, w: TW - 0.22, h: 0.48, fontSize: 12.5, color: textDark, fontFace: FONT, wrap: true });
    });

    if (img) {
      slide.addImage({ data: img, x: IX, y: IY, w: IW, h: IH });
      slide.addShape(pptx.ShapeType.rect, { x: IX, y: IY, w: IW, h: IH, fill: { color: white, alpha: 92 } });
      slide.addShape(pptx.ShapeType.rect, { x: IX - 0.08, y: IY, w: 0.08, h: IH, fill: { color: accent } });
      slide.addShape(pptx.ShapeType.rect, { x: IX, y: IY + IH - 0.05, w: IW, h: 0.05, fill: { color: secondary } });
    } else if (useImg) {
      addPH(slide, IX, IY, IW, IH, 'Add Image Here');
    }
    addFooter(slide, sd.slideNumber);
  };

  const buildVisual = (sd) => {
    const slide = pptx.addSlide();
    const img = pickImg(sd);
    if (img) {
      slide.addImage({ data: img, x: 0, y: 0, w: '100%', h: '100%' });
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 3.45, w: 10, h: 2.18, fill: { color: darkColor, alpha: 28 } });
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.1, fill: { color: accent } });
    } else {
      slide.background = { fill: darkColor };
      addPH(slide, 0.4, 0.2, 9.2, 3.05, 'Visual / Decor Image');
    }
    addLogo(slide);
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 3.45, w: 0.12, h: 2.18, fill: { color: accent } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.18, y: 3.45, w: 0.04, h: 2.18, fill: { color: secondary, alpha: 70 } });
    slide.addText((sd.title || '').toUpperCase(), { x: 0.28, y: 3.55, w: 9.0, h: 0.75, fontSize: 28, bold: true, color: white, fontFace: FONT, charSpacing: 1 });
    if (sd.subtitle) slide.addText(sd.subtitle, { x: 0.28, y: 4.35, w: 7.5, h: 0.42, fontSize: 13, color: 'DDDDDD', fontFace: FONT, italic: true });
    addFooter(slide, sd.slideNumber);
  };

  const buildTimeline = (sd) => {
    const slide = pptx.addSlide();
    slide.background = { fill: bgColor };
    addHeader(slide, sd.title);
    if (sd.subtitle) slide.addText(sd.subtitle, { x: ML, y: CTY, w: 9.3, h: 0.35, fontSize: 12, color: accent, italic: true, fontFace: FONT, align: 'center' });
    const pts = (sd.bulletPoints || []).slice(0, 6);
    const cols = pts.length || 1;
    const colW = 9.2 / cols;
    const lineY = 2.28;
    slide.addShape(pptx.ShapeType.rect, { x: 0.6, y: lineY + 0.37, w: 8.8, h: 0.05, fill: { color: accent, alpha: 55 } });
    pts.forEach((pt, i) => {
      const cx = 0.4 + i * colW + colW / 2;
      slide.addShape(pptx.ShapeType.ellipse, { x: cx - 0.4, y: lineY, w: 0.8, h: 0.8, fill: { color: accent } });
      slide.addShape(pptx.ShapeType.ellipse, { x: cx - 0.35, y: lineY + 0.05, w: 0.7, h: 0.7, fill: { color: secondary, alpha: 70 } });
      slide.addText(String(i + 1), { x: cx - 0.4, y: lineY + 0.07, w: 0.8, h: 0.66, fontSize: 18, bold: true, color: white, align: 'center', fontFace: FONT });
      const parts = pt.split(':');
      slide.addText(parts[0].trim(), { x: cx - colW / 2 + 0.05, y: lineY + 0.92, w: colW - 0.1, h: 0.44, fontSize: 10.5, bold: true, color: textDark, align: 'center', fontFace: FONT });
      if (parts[1]) slide.addText(parts[1].trim(), { x: cx - colW / 2 + 0.05, y: lineY + 1.38, w: colW - 0.1, h: 0.9, fontSize: 9.5, color: textGrey, align: 'center', fontFace: FONT, wrap: true });
    });
    if (sd.bodyText) slide.addText(sd.bodyText, { x: 0.5, y: 4.55, w: 9, h: 0.55, fontSize: 11, color: textGrey, fontFace: FONT, align: 'center' });
    addFooter(slide, sd.slideNumber);
  };

  const buildHighlight = (sd) => {
    const slide = pptx.addSlide();
    slide.background = { fill: bgColor };
    addHeader(slide, sd.title);
    if (sd.subtitle) slide.addText(sd.subtitle, { x: ML, y: CTY, w: 9.3, h: 0.35, fontSize: 12, color: accent, italic: true, fontFace: FONT });
    const icons = ['\u2726', '\u25C6', '\u2605', '\u25CF', '\u25B2'];
    const pts = (sd.bulletPoints || []).slice(0, 3);
    const colW = 9.2 / pts.length;
    pts.forEach((pt, i) => {
      const cx = 0.4 + i * colW;
      slide.addShape(pptx.ShapeType.rect, { x: cx + 0.1, y: 1.48, w: colW - 0.25, h: 3.6, fill: { color: white }, line: { color: 'DDD9D0', width: 0.5 } });
      slide.addShape(pptx.ShapeType.rect, { x: cx + 0.1, y: 1.48, w: colW - 0.25, h: 0.08, fill: { color: accent } });
      slide.addShape(pptx.ShapeType.rect, { x: cx + 0.1, y: 1.56, w: colW - 0.25, h: 0.04, fill: { color: secondary } });
      slide.addShape(pptx.ShapeType.ellipse, { x: cx + colW / 2 - 0.38, y: 1.7, w: 0.76, h: 0.76, fill: { color: bgColor } });
      slide.addText(icons[i % icons.length], { x: cx + colW / 2 - 0.38, y: 1.76, w: 0.76, h: 0.64, fontSize: 22, color: accent, align: 'center' });
      const parts = pt.split(':');
      slide.addText(parts[0].trim(), { x: cx + 0.2, y: 2.55, w: colW - 0.45, h: 0.45, fontSize: 12.5, bold: true, color: textDark, align: 'center', fontFace: FONT });
      if (parts[1]) slide.addText(parts[1].trim(), { x: cx + 0.2, y: 3.05, w: colW - 0.45, h: 1.7, fontSize: 11, color: textGrey, align: 'center', fontFace: FONT, wrap: true });
      slide.addShape(pptx.ShapeType.rect, { x: cx + colW / 2 - 0.5, y: 4.85, w: 1.0, h: 0.04, fill: { color: secondary } });
    });
    addFooter(slide, sd.slideNumber);
  };

  const buildClosing = (sd) => {
    const slide = pptx.addSlide();
    const img = pickImg(sd);
    if (img) {
      slide.addImage({ data: img, x: 0, y: 0, w: '100%', h: '100%' });
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 5.63, fill: { color: darkColor, alpha: 38 } });
    } else {
      slide.background = { fill: darkColor };
      slide.addShape(pptx.ShapeType.rect, { x: 2.5, y: 0, w: 5, h: 5.63, fill: { color: accent, alpha: 90 } });
    }
    if (logoB64) slide.addImage({ data: logoB64, x: 4.25, y: 0.28, w: 1.5, h: 0.72 });
    slide.addShape(pptx.ShapeType.rect, { x: 2.8, y: 1.52, w: 4.4, h: 0.05, fill: { color: accent } });
    slide.addShape(pptx.ShapeType.rect, { x: 2.8, y: 1.6, w: 4.4, h: 0.03, fill: { color: secondary, alpha: 70 } });
    slide.addShape(pptx.ShapeType.rect, { x: 2.8, y: 4.08, w: 4.4, h: 0.05, fill: { color: accent } });
    slide.addShape(pptx.ShapeType.rect, { x: 2.8, y: 4.02, w: 4.4, h: 0.03, fill: { color: secondary, alpha: 70 } });
    slide.addText((sd.title || 'Thank You').toUpperCase(), { x: 0.5, y: 1.75, w: 9, h: 1.4, fontSize: 48, bold: true, color: white, align: 'center', fontFace: FONT, charSpacing: 2 });
    if (sd.subtitle) slide.addText(sd.subtitle, { x: 1, y: 3.28, w: 8, h: 0.55, fontSize: 15, color: 'CCCCCC', align: 'center', fontFace: FONT, italic: true });
    slide.addText((eventData.clientName || '').toUpperCase(), { x: 0.5, y: 4.22, w: 9, h: 0.4, fontSize: 13, color: accent, align: 'center', fontFace: FONT, bold: true, charSpacing: 2 });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.45, w: 10, h: 0.18, fill: { color: accent } });
  };

  for (const sd of slides) {
    const t = (sd.slideType || 'content').toLowerCase();
    const ni = sd.includeImage !== false;
    if (t === 'cover') buildCover(sd);
    else if (t === 'visual') buildVisual(sd);
    else if (t === 'timeline') buildTimeline(sd);
    else if (t === 'highlight') buildHighlight(sd);
    else if (t === 'closing') buildClosing(sd);
    else buildContent(sd, ni);
  }

  return pptx.write({ outputType: 'nodebuffer' });
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Event PPTX service listening on port ${PORT}`));
