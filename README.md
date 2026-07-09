# Event PPTX Generation Service (v2 - JSON based)

Builds your branded event PowerPoint outside n8n (n8n Cloud's Code node can't use pptxgenjs/sharp).

## Update your existing Render deployment

You already deployed this once. To update it with the new version:
1. Go to your GitHub repo (`event-pptx-service`).
2. Open `server.js` in the repo, click the pencil (Edit) icon.
3. Delete everything, paste in the new `server.js` content from this folder. Commit.
4. Do the same for `package.json`.
5. Render will auto-detect the GitHub change and redeploy automatically (watch the Events/Logs tab).

No changes needed to your Environment Variables (API_KEY stays the same).

## API

POST to `https://YOUR-URL.onrender.com/generate-pptx` with:
- Header: `x-api-key: <your API_KEY>`
- Header: `Content-Type: application/json`
- Body (JSON):
```json
{
  "slides": [...],
  "eventData": {...},
  "accentColor": "#...", "secondaryColor": "#...", "bgColor": "#...", "darkColor": "#...",
  "logoBase64": "data:image/png;base64,....",
  "images": ["data:image/jpeg;base64,....", "..."]
}
```
Response: binary `.pptx` file.
