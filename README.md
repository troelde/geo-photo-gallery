# Geo Photo Gallery

A static, client-side-only web app that displays your publicly shared OneDrive photo folder as a date-grouped gallery and an interactive GPS map (photo thumbnails plotted on OpenStreetMap / OpenTopoMap / satellite imagery).

**Live demo:** https://troelde.github.io/geo-photo-gallery/

---

## Quick Start

### 1 — Share your OneDrive folder

1. Open OneDrive → right-click your photos folder → **Share**
2. Set link type to **"Anyone with the link can view"**
3. Copy the share URL (looks like `https://1drv.ms/f/...`)

### 2 — Register a free Azure App (~5 minutes)

Microsoft Graph requires a signed-in user to read a share — even "Anyone with the link" shares — so this step is mandatory, not optional.

1. Go to <https://portal.azure.com> → **App registrations** → **New registration**
2. Name: `geo-photo-gallery` (or anything)
3. Supported account types: **"Personal Microsoft accounts only"**
4. Redirect URI: **Single-page application (SPA)** → `http://localhost:8080/`
   - Add your production URL too if you deploy (e.g. `https://<you>.github.io/geo-photo-gallery/`)
5. Click **Register** → copy the **Application (client) ID**

### 3 — Fill in `config.js`

```js
const CONFIG = {
  shareUrl: 'https://1drv.ms/f/...YOUR_ACTUAL_LINK',
  clientId: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
  map: {
    defaultCenter: [20, 0],
    defaultZoom: 2,
  },
};
```

### 4 — Run locally

You need a local HTTP server (required for MSAL's redirect-based sign-in — `file://` won't work).

```bash
# Python
python3 -m http.server 8080

# Node.js (npx, no install)
npx serve . -p 8080

# VS Code: install the "Live Server" extension and click "Go Live"
```

Then open **http://localhost:8080** and click **Sign in to Microsoft**.

---

## Authentication

Sign-in uses [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) with the **redirect flow** (full-page navigation), not a popup — this avoids a class of bugs where popup windows can hang open indefinitely after a successful sign-in due to browser timer throttling.

1. Click **Sign in to Microsoft** → you're redirected to a Microsoft login page
2. Authenticate with the Microsoft account that owns (or has access to) the OneDrive share
3. You're redirected back and your photos load automatically
4. The session is cached in `sessionStorage`, so you won't need to sign in again until you close the tab

---

## Features

| Feature | Details |
|---------|---------|
| Date-grouped gallery | Photos grouped by day taken, sorted chronologically (oldest first); undated photos last |
| Lightbox | Click any photo for a full-resolution view, fetched on demand from Microsoft Graph, with prev/next navigation |
| Jump to map | Photos with GPS data show a "📍 View on map" button in the lightbox that switches to the Map tab centered on that photo's pin |
| Photo descriptions | OneDrive file descriptions are shown as an overlay caption in the lightbox and as a hover tooltip in the gallery grid |
| Gallery & per-day descriptions | Optional Markdown files in a `metadata` subfolder are rendered above the gallery and/or under each date heading (see below) |
| Map view | GPS-tagged photos plotted as thumbnail pins, with a layer switcher for Topographic (OpenTopoMap), Streets (OSM), and Satellite (Esri) |
| No-GPS handling | A legend shows how many photos lack GPS data; those are still browsable in the Gallery |
| YAML sidecar fallback | A `<photo filename>.yaml` file next to a photo can supply GPS coordinates and/or a description when OneDrive doesn't already have them (see below) |
| Sticky header/tabs | The title bar and Gallery/Map tabs stay pinned while scrolling |
| Pagination | Automatically fetches all pages from Graph (>200 photos) |
| No build tools | Plain HTML/CSS/JS, zero dependencies to install |

---

## Gallery & Per-Day Descriptions

You can add optional Markdown files to a `metadata` subfolder inside your shared OneDrive folder to add narrative text to the gallery:

| File | Where it's rendered |
|------|----------------------|
| `metadata/description.md` | Above the whole gallery, as an introduction/overview |
| `metadata/YYYYMMDD-description.md` | Under that day's date heading in the gallery (e.g. `metadata/20250712-description.md` for July 12, 2025) |

Both support standard Markdown (headings, bold/italic, links, lists, etc.), rendered client-side with [marked.js](https://marked.js.org/). Links always open in a new tab so readers don't navigate away from the gallery. Files are optional — omit either or both if you don't need them.

---

## GPS & Description Fallback via YAML Sidecar Files

If a photo has no GPS location in its EXIF data and/or no description set in OneDrive, you can supply either manually by adding a YAML file right next to it in the same shared folder, named `<photo filename>.yaml` (e.g. `IMG_1234.jpg.yaml` for `IMG_1234.jpg`):

```yaml
position:
  lat: 78.22334
  long: 15.6482
description: Arrival day in Longyearbyen
```

Both fields are optional and independent — include only what you need. The sidecar is only used to fill gaps: it never overrides a photo's real EXIF GPS location or an existing OneDrive description. This is handy for scanned photos, screenshots, or camera gear that doesn't record GPS/descriptions.

---

## Deploying

Any static hosting works: **GitHub Pages**, **Netlify**, **Vercel**, **Azure Static Web Apps**, etc. This repo is deployed via GitHub Pages, serving directly from the `main` branch root — no build step needed.

After deploying, add your production URL as an additional **Redirect URI** in your Azure App Registration (Authentication blade → SPA platform), exactly matching the deployed URL (including trailing slash), alongside `http://localhost:8080/`.

> ⚠️ Because `config.js` is committed to the repo, your OneDrive share URL and Azure client ID become publicly visible if the repo is public. The client ID is not a secret for public client SPA apps, but only commit a share URL you're comfortable exposing.

---

## Viewing a Different Share Without Editing config.js

You can override the configured share URL at runtime with a `?shareUrl=` query parameter, e.g.:

```
https://troelde.github.io/geo-photo-gallery/?shareUrl=https://1drv.ms/f/...OTHER_LINK
```

This is handy for sharing links to a specific album without redeploying. If the parameter is absent, `config.js`'s `shareUrl` is used as the default.

---

## File Structure

```
geo-photo-gallery/
├── index.html    — App shell + HTML structure
├── app.js        — MSAL auth, OneDrive Graph API, gallery + map logic
├── style.css     — Dark-theme responsive styles
├── config.js     — Your share URL and Azure client ID (edit this)
└── README.md     — This file
```

---

## Known Gotchas

- **Browser caching**: GitHub Pages serves static assets with `Cache-Control: max-age=600`. After pushing an update, do a hard refresh (Ctrl+Shift+R) or use a private/incognito window if you don't see changes immediately.
- **Redirect URI mismatch**: `redirectUri` is computed as `window.location.origin + window.location.pathname` at runtime — it must exactly match an entry in your Azure App Registration's Redirect URIs list, or sign-in will fail.
