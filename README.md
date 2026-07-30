# Geo Photo Gallery

A static web app that displays your publicly shared OneDrive photo folder as a responsive gallery with an interactive map view (GPS-tagged photos plotted on OpenStreetMap).

---

## Quick Start

### 1 — Share your OneDrive folder

1. Open OneDrive → right-click your photos folder → **Share**
2. Set link type to **"Anyone with the link can view"**
3. Copy the share URL (looks like `https://1drv.ms/f/s!XXXXXXX`)

### 2 — Paste the share URL into `config.js`

```js
shareUrl: 'https://1drv.ms/f/s!YOUR_ACTUAL_LINK',
```

### 3 — Run locally

You need a local HTTP server (required for MSAL redirect URIs — `file://` won't work).

```bash
# Python
python3 -m http.server 8080

# Node.js (npx, no install)
npx serve . -p 8080

# VS Code: install the "Live Server" extension and click "Go Live"
```

Then open **http://localhost:8080** in your browser.

---

## Authentication

The app first tries to access your OneDrive share **without any login**.
If Microsoft returns a `401 Unauthorized`, you need to register a free Azure app:

### Register a free Azure App (~5 minutes)

1. Go to <https://portal.azure.com> → **App registrations** → **New registration**
2. Name: `geo-photo-gallery` (or anything)
3. Supported account types: **"Personal Microsoft accounts only"**
4. Redirect URI: **Single-page application (SPA)** → `http://localhost:8080`
   - Add your production URL later if you deploy
5. Click **Register** → copy the **Application (client) ID**
6. Paste it into `config.js`:

```js
clientId: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
```

7. Reload the app → click **Sign in to Microsoft** → authenticate with your Microsoft account (the OneDrive owner)
8. Your photos will load. The session is cached, so you won't need to sign in every time.

---

## Features

| Feature | Details |
|---------|---------|
| Photo grid | Responsive masonry-style grid, lazy-loaded thumbnails |
| Lightbox | Click any photo to view large version with metadata |
| Map view | GPS-tagged photos plotted as thumbnail pins on OpenStreetMap |
| Pagination | Automatically fetches all pages (>200 photos) |
| No build tools | Plain HTML/CSS/JS, zero dependencies to install |

---

## Deploying

Any static hosting works: **GitHub Pages**, **Netlify**, **Vercel**, **Azure Static Web Apps**, etc.

After deploying, add your production URL as an additional SPA Redirect URI in your Azure app registration.

---

## File Structure

```
geo-photo-gallery/
├── index.html    — App shell + HTML structure
├── app.js        — OneDrive Graph API + map + gallery logic
├── style.css     — Dark-theme responsive styles
├── config.js     — Your share URL and Azure client ID (edit this)
└── README.md     — This file
```
