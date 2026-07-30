// ============================================================
// Geo Photo Gallery — Configuration
// Edit this file to connect to your OneDrive share.
// ============================================================

const CONFIG = {
  // Paste your OneDrive "Anyone with the link" share URL here.
  // How to get it:
  //   OneDrive → right-click your photos folder → Share
  //   → "Anyone with the link can view" → Copy link
  shareUrl: 'https://1drv.ms/f/c/99ad336b2325601b/IgBWQR1Nwb5eRZ09VnaZO-ZzAVRlkuoNDh1EP8GDBcI2-9A',

  // Azure App Client ID — needed only if the anonymous fetch fails (likely).
  // How to get one (free, ~5 min):
  //   1. Go to https://portal.azure.com → "App registrations" → New registration
  //   2. Name: anything (e.g. "geo-photo-gallery")
  //   3. Supported account types: "Personal Microsoft accounts only"
  //   4. Redirect URI → Single-page application (SPA) → http://localhost:8080
  //      (add your deployed URL later if needed)
  //   5. Copy the "Application (client) ID" here
  clientId: 'f37e0195-c8c1-4473-afa5-b30a10783b8c',

  // Leaflet map defaults
  map: {
    defaultCenter: [20, 0],
    defaultZoom: 2,
  },
};
