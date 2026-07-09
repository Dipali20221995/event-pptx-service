# Event PPTX Generation Service (v3 - now with image compression)

Builds your branded event PowerPoint AND compresses oversized photos, both
outside n8n (n8n Cloud's Starter plan has only 320MB RAM per execution,
which is too tight to decode/resize large phone photos).

## Update your existing Render deployment

1. Go to your GitHub repo (`event-pptx-service`).
2. Open `server.js`, click the pencil (Edit) icon.
3. Select all, delete, paste in the new `server.js` content from this folder. Commit.
4. Render auto-redeploys within ~1-2 minutes (watch the Events tab).

No changes to `package.json` or your Environment Variables needed.

## New endpoint: /compress-image

POST any image (any size) as raw bytes, get back a small compressed JPEG.

- URL: `https://YOUR-URL.onrender.com/compress-image?quality=70&maxWidth=1280&maxHeight=1280`
- Header: `x-api-key: <your API_KEY>`
- Header: `Content-Type: application/octet-stream` (or the image's real mimetype)
- Body: raw image bytes
- Response: compressed JPEG bytes

Tested: a 10.5MB photo compresses to ~490KB at these settings.

## Existing /generate-pptx endpoint

Unchanged from before — still builds the branded PPTX from JSON slide data + base64 images.
