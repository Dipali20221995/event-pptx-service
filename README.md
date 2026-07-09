# Event PPTX Generation Service

This tiny web service builds your branded event PowerPoint (pptxgenjs + sharp)
outside of n8n, since n8n Cloud's Code node can't use those packages.

## Deploy for free on Render.com (no credit card, no billing)

1. Go to https://github.com and create a new **empty** repository (e.g. `event-pptx-service`).
2. Click **"Add file" → "Upload files"** in that repo and upload `server.js` and `package.json` from this folder.
3. Go to https://render.com and sign up (free, no card required).
4. Click **New +** → **Web Service** → connect your GitHub account → select the repo you just created.
5. Settings:
   - **Runtime**: Node
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Instance Type**: Free
6. Under **Environment Variables**, add:
   - `API_KEY` = any secret string you make up (e.g. `mySuperSecret123`) — this stops random people from hitting your endpoint.
7. Click **Create Web Service**. Wait ~2 minutes for the first deploy.
8. Copy your live URL (looks like `https://event-pptx-service.onrender.com`).

Note: Render's free tier spins the service down after 15 minutes of no traffic.
The first request after idle takes ~30-50 seconds to "wake up" — totally fine for
a form-based workflow that runs occasionally.

## Using it from n8n

POST to `https://YOUR-URL.onrender.com/generate-pptx` as `multipart/form-data`:
- `payload` (text field): JSON string with `slides`, `eventData`, `accentColor`, `secondaryColor`, `bgColor`, `darkColor`
- `logo` (file): the company logo
- `images` (file, repeat for each): reference images

Header: `x-api-key: <the API_KEY you set>`

Response: binary `.pptx` file.
