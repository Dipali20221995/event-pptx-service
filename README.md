# Event PPTX Generation Service

This tiny web service builds your branded event PowerPoint (pptxgenjs + sharp)
outside of n8n, since n8n Cloud's Code node can't use those packages. It also
handles image compression for the AI vision-analysis step earlier in the
workflow.

## Deploy for free on Render.com (no credit card, no billing)

1. Go to https://github.com and create (or reuse) a repository, e.g. `event-pptx-service`.
2. Upload all three files from this folder into the repo root: `server.js`, `package.json`, `README.md`.
3. Go to https://render.com and sign up (free, no card required).
4. Click **New +** → **Web Service** → connect your GitHub account → select the repo.
5. Settings:
   - **Runtime**: Node
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Instance Type**: Free
6. Under **Environment Variables**, add:
   - `API_KEY` = a secret string (this must exactly match the `x-api-key` header value n8n sends, e.g. `Ppt2026secret`).
7. Click **Create Web Service**. Wait ~2 minutes for the first deploy.
8. Copy your live URL (looks like `https://event-pptx-service.onrender.com`).
9. Check the **Logs** tab — you should see `Event PPTX service listening on port ...` with no red errors.

If you ever change `server.js` or `package.json` again: replace the file(s) on
GitHub via the web editor, commit to `main`, and Render will auto-redeploy.
Watch the Logs tab each time to confirm a clean start before testing from n8n.

Note: Render's free tier spins the service down after 15 minutes of no traffic.
The first request after idle takes ~30-50 seconds to "wake up" — totally fine
for a form-based workflow that runs occasionally, as long as the calling side
(n8n) uses a generous timeout and/or retry, which this workflow's HTTP nodes
already do.

## Endpoints

### `POST /compress-image?quality=70&maxWidth=1280&maxHeight=1280`
Used by n8n before AI vision analysis, to shrink each uploaded reference photo.
- Request body: the raw image bytes (n8n sends this as `contentType: binaryData`,
  i.e. a raw POST body, **not** multipart/form-data).
- Query params: `quality` (JPEG quality, default 80), `maxWidth`, `maxHeight`
  (default 1280 each) — image is resized to fit inside these bounds without
  upscaling.
- Header: `x-api-key: <the API_KEY you set>`
- Response: binary JPEG file.

### `POST /generate-pptx`
Used at the end of the workflow to build the final PowerPoint.
- Request format: `multipart/form-data`
  - `payload` (text field): JSON string with `slides` (each slide may include
    an `imageIndex` picking which uploaded reference image — by position —
    belongs on that slide; omit or leave `null` for slides that should have
    no image), `eventData`, `accentColor`, `secondaryColor`, `bgColor`, `darkColor`.
  - `logo` (file, optional): the company logo. If omitted, slides show a
    blank "LOGO" placeholder instead of a real logo.
  - `images` (file, repeat once per reference image, in the order your
    `imageIndex` values refer to): full-resolution originals are fine — this
    service crops (16:9 cover fit) and brightens/adjusts them itself via `sharp`.
- Header: `x-api-key: <the API_KEY you set>`
- Response: binary `.pptx` file.
