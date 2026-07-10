const express = require('express');
const multer = require('multer');
const PptxGenJS = require('pptxgenjs');
const sharp = require('sharp');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

const API_KEY = process.env.API_KEY || '';

app.get('/', (req, res) => res.send('Event PPTX service is running.'));

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
  const accent    = (inputData.accentColor    || '#B38E58').replace('#', '');
  const secondary = (inputData.secondaryColor || '#D4B896').replace('#', '');
  const bgColor   = (inputData.bgColor        || '#FBF7F2').replace('#', '');
  const darkColor = (inputData.darkColor      || '#3D2820').replace('#', '');
  const white     = 'FFFFFF';
  const textDark  = '2C2620';
  const textGrey  = '746B60';
  const hairline  = 'DDD5C8';

  const DISPLAY = 'Georgia';
  const BODY    = 'Calibri';

  const ML     = 0.55;
  const MR     = 9.45;
  const FTR_Y  = 5.32;
  const LOGO_W = 0.62;
  const LOGO_H = 0.62;

  const refImgs = [];
  let logoB64 = null;

  if (logoFile) {
    try {
      const lb = await sharp(logoFile.buffer)
        .resize(240, 240, { fit: 'inside', background: { r: 255, g: 255, b: 255, alpha: 0 } })
        .png()
        .toBuffer();
      logoB64 = 'data:image/png;base64,' + lb.toString('base64');
    } catch (e) { /* skip bad logo */ }
  }

  for (const f of refFiles) {
    try {
      const rb = await sharp(f.buffer)
        .resize(1600, 900, { fit: 'cover', position: 'attention' })
        .modulate({ brightness: 1.08, saturation: 0.9 })
        .jpeg({ quality: 92 })
        .toBuffer();
      refImgs.push('data:image/jpeg;base64,' + rb.toString('base64'));
    } catch (e) { /* skip bad image */ }
  }

  const pickImg = (sd) => {
    const idx = sd && sd.imageIndex;
    if (typeof idx === 'number' && idx >= 0 && idx < refImgs.length) {
      return refImgs[idx];
    }
    return null;
  };

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

  const ICON_MAP = [
    [/client|about/i, '\u25EF'],
    [/overview/i, '\u2261'],
    [/theme|design|concept/i, '\u25C8'],
    [/venue/i, '\u25C7'],
    [/schedule|timeline/i, '\u29C9'],
    [/guest|experience/i, '\u25CE'],
    [/stage|decor/i, '\u2727'],
    [/hospitality|catering/i, '\u25C9'],
    [/logistics|operations/i, '\u2318'],
    [/why choose|why us/i, '\u2606'],
    [/thank/i, '\u2661']
  ];
  const pickIcon = (title) => {
    const t = String(title || '');
    for (const [re, icon] of ICON_MAP) {
      if (re.test(t)) return icon;
    }
    return '\u25C7';
  };

  const diamondDivider = (slide, cx, y, w, color) => {
    const c = color || accent;
    const gap = 0.16;
    const half = (w - gap * 2) / 2;
    slide.addShape(pptx.ShapeType.line, { x: cx - w / 2, y, w: half, h: 0, line: { color: c, width: 0.75 } });
    slide.addShape(pptx.ShapeType.rect, { x: cx - 0.045, y: y - 0.045, w: 0.09, h: 0.09, rotate: 45, fill: { color: c } });
    slide.addShape(pptx.ShapeType.line, { x: cx + gap, y, w: half, h: 0, line: { color: c, width: 0.75 } });
  };

  const cornerBrackets = (slide, color) => {
    const c = color || accent, len = 0.42, o = 0.32, w = 0.75;
    slide.addShape(pptx.ShapeType.line, { x: o, y: o, w: len, h: 0, line: { color: c, width: w } });
    slide.addShape(pptx.ShapeType.line, { x: o, y: o, w: 0, h: len, line: { color: c, width: w } });
    slide.addShape(pptx.ShapeType.line, { x: 10 - o - len, y: o, w: len, h: 0, line: { color: c, width: w } });
    slide.addShape(pptx.ShapeType.line, { x: 10 - o, y: o, w: 0, h: len, line: { color: c, width: w } });
    slide.addShape(pptx.ShapeType.line, { x: o, y: 5.63 - o, w: len, h: 0, line: { color: c, width: w } });
    slide.addShape(pptx.ShapeType.line, { x: o, y: 5.63 - o - len, w: 0, h: len, line: { color: c, width: w } });
    slide.addShape(pptx.ShapeType.line, { x: 10 - o - len, y: 5.63 - o, w: len, h: 0, line: { color: c, width: w } });
    slide.addShape(pptx.ShapeType.line, { x: 10 - o, y: 5.63 - o - len, w: 0, h: len, line: { color: c, width: w } });
  };

  const framedImage = (slide, img, x, y, w, h) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x - 0.06, y: y - 0.06, w: w + 0.12, h: h + 0.12,
      rectRadius: 0.06, fill: { color: white },
      line: { color: hairline, width: 0.75 },
      shadow: { type: 'outer', color: '000000', opacity: 0.18, blur: 10, offset: 3, angle: 90 }
    });
    slide.addImage({ data: img, x, y, w, h, rounding: false });
  };

  const addLogo = (slide, x, y, color) => {
    if (logoB64) {
      slide.addImage({ data: logoB64, x, y, w: LOGO_W, h: LOGO_H });
    } else {
      slide.addShape(pptx.ShapeType.ellipse, { x, y, w: LOGO_W, h: LOGO_H, fill: { color: 'FFFFFF', alpha: 6 }, line: { color: color || accent, width: 0.75 } });
      slide.addText('LOGO', { x, y, w: LOGO_W, h: LOGO_H, fontSize: 7, color: color || accent, align: 'center', fontFace: BODY, bold: true, valign: 'middle' });
    }
  };

  const addFooter = (slide, num) => {
    slide.addShape(pptx.ShapeType.line, { x: ML, y: FTR_Y, w: MR - ML, h: 0, line: { color: hairline, width: 0.75 } });
    slide.addText((eventData.clientName || '').toUpperCase(), { x: ML, y: FTR_Y + 0.06, w: 4, h: 0.2, fontSize: 7.5, color: textGrey, fontFace: BODY, charSpacing: 1 });
    slide.addText((eventData.eventName || '').toUpperCase(), { x: 3, y: FTR_Y + 0.06, w: 4, h: 0.2, fontSize: 7.5, color: textGrey, fontFace: BODY, align: 'center', charSpacing: 1 });
    if (num) slide.addText(String(num).padStart(2, '0'), { x: MR - 0.5, y: FTR_Y + 0.06, w: 0.5, h: 0.2, fontSize: 7.5, color: textGrey, align: 'right', fontFace: BODY, charSpacing: 1 });
  };

  const addHeading = (slide, title, eyebrow) => {
    const icon = pickIcon(title);
    slide.addText(icon + '   ' + (eyebrow || 'SECTION').toUpperCase(), {
      x: ML, y: 0.42, w: MR - ML, h: 0.32, fontSize: 11, color: accent, fontFace: BODY, charSpacing: 3, bold: true
    });
    slide.addText(title || '', {
      x: ML, y: 0.74, w: MR - ML, h: 0.72, fontSize: 30, color: textDark, fontFace: DISPLAY, italic: true, valign: 'top'
    });
    diamondDivider(slide, 5.0, 1.52, 2.2, accent);
    addLogo(slide, MR - LOGO_W, 0.4, accent);
  };

  const buildCover = (sd) => {
    const slide = pptx.addSlide();
    slide.background = { fill: darkColor };
    cornerBrackets(slide, secondary);

    const img = pickImg(sd);
    if (img) {
      slide.addImage({ data: img, x: 0, y: 0, w: '100%', h: '100%' });
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 5.63, fill: { color: darkColor, alpha: 22 } });
    }

    addLogo(slide, 4.69, 0.42, secondary);
    slide.addText('PRESENTS', { x: 0, y: 1.16, w: 10, h: 0.3, fontSize: 11, color: secondary, align: 'center', fontFace: BODY, charSpacing: 4, bold: true });
    slide.addText((eventData.eventBrief || eventData.theme || '').toString(), { x: 1, y: 1.5, w: 8, h: 0.4, fontSize: 13, color: 'D8CFC2', align: 'center', fontFace: DISPLAY, italic: true });

    slide.addText(eventData.clientName || '', {
      x: 0.4, y: 1.95, w: 9.2, h: 1.15, fontSize: 46, color: white, align: 'center', fontFace: DISPLAY, italic: true, wrap: true, fit: 'shrink'
    });

    diamondDivider(slide, 5.0, 3.18, 2.6, secondary);

    slide.addText((eventData.eventName || sd.title || '').toUpperCase(), { x: 0, y: 3.32, w: 10, h: 0.4, fontSize: 18, color: white, align: 'center', fontFace: BODY, bold: true, charSpacing: 2 });
    if (eventData.theme) {
      slide.addText(eventData.theme, { x: 1, y: 3.72, w: 8, h: 0.35, fontSize: 12, color: secondary, align: 'center', fontFace: DISPLAY, italic: true });
    }

    const boxY = 4.35, boxH = 0.78, boxX = 1.4, boxW = 7.2;
    slide.addShape(pptx.ShapeType.rect, { x: boxX, y: boxY, w: boxW, h: boxH, fill: { color: 'FFFFFF', alpha: 96 }, line: { color: secondary, width: 0.75 } });
    const half = boxW / 2;
    slide.addShape(pptx.ShapeType.line, { x: boxX + half, y: boxY + 0.14, w: 0, h: boxH - 0.28, line: { color: secondary, width: 0.5 } });
    slide.addText('DATE', { x: boxX, y: boxY + 0.1, w: half, h: 0.2, fontSize: 8.5, color: secondary, align: 'center', fontFace: BODY, charSpacing: 2, bold: true });
    slide.addText(eventData.eventDate || '', { x: boxX, y: boxY + 0.32, w: half, h: 0.36, fontSize: 13, color: white, align: 'center', fontFace: DISPLAY, italic: true });
    slide.addText('VENUE', { x: boxX + half, y: boxY + 0.1, w: half, h: 0.2, fontSize: 8.5, color: secondary, align: 'center', fontFace: BODY, charSpacing: 2, bold: true });
    slide.addText(eventData.venue || '', { x: boxX + half, y: boxY + 0.32, w: half, h: 0.36, fontSize: 13, color: white, align: 'center', fontFace: DISPLAY, italic: true });
  };

  const buildContent = (sd, useImg) => {
    const slide = pptx.addSlide();
    slide.background = { fill: bgColor };
    addHeading(slide, sd.title, (sd.slideType || 'section'));

    const hasImg = useImg && pickImg(sd);
    const TW = hasImg ? 4.9 : MR - ML;
    const IX = 5.75, IY = 1.75, IW = 3.7, IH = 3.35;

    let y = 1.85;
    if (sd.subtitle) {
      slide.addText(sd.subtitle, { x: ML, y, w: TW, h: 0.34, fontSize: 12.5, color: accent, italic: true, fontFace: DISPLAY });
      y += 0.46;
    }
    if (sd.bodyText) {
      slide.addText(sd.bodyText, { x: ML, y, w: TW, h: 0.68, fontSize: 11.5, color: textGrey, fontFace: BODY, wrap: true, lineSpacingMultiple: 1.25 });
      y += 0.8;
    }

    (sd.bulletPoints || []).forEach((pt) => {
      slide.addShape(pptx.ShapeType.rect, { x: ML, y: y + 0.09, w: 0.07, h: 0.07, rotate: 45, fill: { color: accent } });
      slide.addText(pt, { x: ML + 0.22, y, w: TW - 0.3, h: 0.46, fontSize: 12, color: textDark, fontFace: BODY, wrap: true, lineSpacingMultiple: 1.1 });
      y += 0.5;
    });

    if (hasImg) {
      framedImage(slide, pickImg(sd), IX, IY, IW, IH);
    } else if (useImg) {
      slide.addShape(pptx.ShapeType.roundRect, { x: IX, y: IY, w: IW, h: IH, rectRadius: 0.06, fill: { color: 'F2EEE6' }, line: { color: hairline, width: 0.75, dashType: 'dash' } });
      slide.addText('[ space reserved for image ]', { x: IX, y: IY + IH / 2 - 0.2, w: IW, h: 0.4, fontSize: 10, color: accent, align: 'center', fontFace: DISPLAY, italic: true });
    }
    addFooter(slide, sd.slideNumber);
  };

  const buildVisual = (sd) => {
    const slide = pptx.addSlide();
    const img = pickImg(sd);
    if (img) {
      slide.addImage({ data: img, x: 0, y: 0, w: '100%', h: '100%' });
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 3.7, w: 10, h: 1.93, fill: { color: darkColor, alpha: 30 } });
    } else {
      slide.background = { fill: darkColor };
    }
    addLogo(slide, MR - LOGO_W, 0.35, secondary);
    diamondDivider(slide, 1.7, 3.95, 1.3, secondary);
    slide.addText((sd.title || '').toUpperCase(), { x: 0.55, y: 4.08, w: 8.5, h: 0.6, fontSize: 26, color: white, fontFace: DISPLAY, italic: true });
    if (sd.subtitle) slide.addText(sd.subtitle, { x: 0.55, y: 4.68, w: 8, h: 0.4, fontSize: 12.5, color: 'E6DFD3', fontFace: BODY, italic: false, charSpacing: 1 });
    addFooter(slide, sd.slideNumber);
  };

  const buildTimeline = (sd) => {
    const slide = pptx.addSlide();
    slide.background = { fill: bgColor };
    addHeading(slide, sd.title, 'timeline');
    if (sd.subtitle) slide.addText(sd.subtitle, { x: ML, y: 1.82, w: MR - ML, h: 0.32, fontSize: 12.5, color: accent, italic: true, fontFace: DISPLAY, align: 'center' });

    const pts = (sd.bulletPoints || []).slice(0, 6);
    const cols = pts.length || 1;
    const colW = (MR - ML) / cols;
    const lineY = 2.85;
    slide.addShape(pptx.ShapeType.line, { x: ML + 0.4, y: lineY, w: (MR - ML) - 0.8, h: 0, line: { color: hairline, width: 1 } });
    pts.forEach((pt, i) => {
      const cx = ML + i * colW + colW / 2;
      slide.addShape(pptx.ShapeType.ellipse, { x: cx - 0.22, y: lineY - 0.22, w: 0.44, h: 0.44, fill: { color: bgColor }, line: { color: accent, width: 1 } });
      slide.addText(String(i + 1).padStart(2, '0'), { x: cx - 0.22, y: lineY - 0.19, w: 0.44, h: 0.38, fontSize: 11, bold: true, color: accent, align: 'center', fontFace: BODY });
      const parts = pt.split(':');
      slide.addText(parts[0].trim(), { x: cx - colW / 2 + 0.06, y: lineY + 0.32, w: colW - 0.12, h: 0.4, fontSize: 12, bold: true, color: textDark, align: 'center', fontFace: DISPLAY, italic: true });
      if (parts[1]) slide.addText(parts[1].trim(), { x: cx - colW / 2 + 0.06, y: lineY + 0.74, w: colW - 0.12, h: 1.0, fontSize: 9.5, color: textGrey, align: 'center', fontFace: BODY, wrap: true, lineSpacingMultiple: 1.15 });
    });
    if (sd.bodyText) slide.addText(sd.bodyText, { x: 0.8, y: 4.7, w: 8.4, h: 0.5, fontSize: 10.5, color: textGrey, fontFace: BODY, align: 'center', italic: true });
    addFooter(slide, sd.slideNumber);
  };

  const buildHighlight = (sd) => {
    const slide = pptx.addSlide();
    slide.background = { fill: bgColor };
    addHeading(slide, sd.title, 'highlights');
    if (sd.subtitle) slide.addText(sd.subtitle, { x: ML, y: 1.82, w: MR - ML, h: 0.32, fontSize: 12.5, color: accent, italic: true, fontFace: DISPLAY, align: 'center' });

    const pts = (sd.bulletPoints || []).slice(0, 3);
    const gap = 0.22;
    const colW = ((MR - ML) - gap * (pts.length - 1)) / pts.length;
    pts.forEach((pt, i) => {
      const cx = ML + i * (colW + gap);
      slide.addShape(pptx.ShapeType.roundRect, {
        x: cx, y: 2.32, w: colW, h: 2.75, rectRadius: 0.08,
        fill: { color: white }, line: { color: hairline, width: 0.75 },
        shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 8, offset: 2, angle: 90 }
      });
      slide.addShape(pptx.ShapeType.ellipse, { x: cx + colW / 2 - 0.3, y: 2.55, w: 0.6, h: 0.6, fill: { color: bgColor }, line: { color: accent, width: 0.75 } });
      slide.addText(pickIcon(pt), { x: cx + colW / 2 - 0.3, y: 2.55, w: 0.6, h: 0.6, fontSize: 18, color: accent, align: 'center', valign: 'middle', fontFace: DISPLAY });
      const parts = pt.split(':');
      slide.addText(parts[0].trim(), { x: cx + 0.15, y: 3.35, w: colW - 0.3, h: 0.42, fontSize: 13, bold: false, color: textDark, align: 'center', fontFace: DISPLAY, italic: true });
      diamondDivider(slide, cx + colW / 2, 3.85, 0.6, secondary);
      if (parts[1]) slide.addText(parts[1].trim(), { x: cx + 0.2, y: 3.98, w: colW - 0.4, h: 1.0, fontSize: 10, color: textGrey, align: 'center', fontFace: BODY, wrap: true, lineSpacingMultiple: 1.15 });
    });
    addFooter(slide, sd.slideNumber);
  };

  const buildClosing = (sd) => {
    const slide = pptx.addSlide();
    slide.background = { fill: darkColor };
    cornerBrackets(slide, secondary);

    addLogo(slide, 4.69, 0.55, secondary);
    slide.addText('WITH LOVE', { x: 0, y: 1.3, w: 10, h: 0.3, fontSize: 11, color: secondary, align: 'center', fontFace: BODY, charSpacing: 4, bold: true });

    slide.addText((sd.title || 'Thank You'), {
      x: 0.5, y: 1.68, w: 9, h: 1.1, fontSize: 54, color: white, align: 'center', fontFace: DISPLAY, italic: true, fit: 'shrink'
    });
    diamondDivider(slide, 5.0, 2.86, 2.2, secondary);
    if (sd.subtitle) slide.addText(sd.subtitle, { x: 1, y: 3.02, w: 8, h: 0.5, fontSize: 14, color: 'D8CFC2', align: 'center', fontFace: DISPLAY, italic: true });

    slide.addText(
      [eventData.eventDate, eventData.venue].filter(Boolean).join('   \u00B7   ').toUpperCase(),
      { x: 0.5, y: 4.55, w: 9, h: 0.35, fontSize: 11, color: secondary, align: 'center', fontFace: BODY, charSpacing: 2, bold: true }
    );
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
