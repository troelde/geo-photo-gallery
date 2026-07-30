// ============================================================
// Geo Photo Gallery — Application Logic
//
// Auth pattern (redirect flow) adapted from Microsoft's official
// sample: ms-identity-javascript-v2 (app/authRedirect.js).
//
// We use full-page redirects instead of popups because popup-based
// auth relies on the opener window polling the popup's URL to detect
// completion and close it — this polling can be silently throttled
// or blocked by some browsers, leaving the popup open indefinitely
// even after a fully successful sign-in. Redirect flow avoids that
// class of problem entirely.
// ============================================================

const GRAPH_API = 'https://graph.microsoft.com/v1.0';

/** MSAL config - mirrors ms-identity-javascript-v2/app/authConfig.js */
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: 'https://login.microsoftonline.com/consumers',
    // Redirect flow lands back on this same app page; handleRedirectPromise()
    // (called at boot) picks up the response after the round-trip.
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

// Request both scopes at sign-in so a single consent screen covers
// everything the app needs (avoids a second redirect round-trip).
const loginRequest = { scopes: ['User.Read', 'Files.Read'] };
const tokenRequest = { scopes: ['Files.Read'] };

let msalInstance = null;
let username = '';
let homeAccountId = null; // more reliable account key than username; used with getAccount()
let leafletMap = null;
let lastMapBounds = null; // re-applied once the map tab becomes visible
let photos = [];
let pendingAccessToken = null; // set when handleRedirectPromise() itself returns a Graph token
let lastAccessToken = null; // cached for on-demand full-res fetches from the lightbox

// ---- Helpers ------------------------------------------------

/** Encode a OneDrive share URL into the Graph API shareId format. */
function encodeShareUrl(url) {
  const b64 = btoa(url)
    .replace(/=+$/, '')
    .replace(/\//g, '_')
    .replace(/\+/g, '-');
  return 'u!' + b64;
}

/** Fetch all image children of a shared folder, following nextLink pages. */
async function fetchShareChildren(token) {
  const shareId = encodeShareUrl(CONFIG.shareUrl);
  const select = [
    'id', 'name', 'description', 'file', 'photo', 'location', 'image',
    'parentReference', '@microsoft.graph.downloadUrl',
  ].join(',');

  const firstPage =
    `${GRAPH_API}/shares/${shareId}/driveItem/children` +
    `?$select=${select}&$expand=thumbnails($select=medium,large)&$top=200`;

  const headers = token ? { Authorization: `Bearer ${token}` } : {};
  const items = [];
  let url = firstPage;

  while (url) {
    const resp = await fetch(url, { headers });
    if (!resp.ok) return { ok: false, status: resp.status };
    const data = await resp.json();
    items.push(...(data.value ?? []));
    url = data['@odata.nextLink'] ?? null;
  }

  return { ok: true, items };
}

// ---- MSAL Auth ------------------------------------------------
// Pattern adapted from ms-identity-javascript-v2/app/authRedirect.js

function initMsal() {
  if (!CONFIG.clientId) return null;
  return new msal.PublicClientApplication(msalConfig);
}

/** Picks up a cached account (if any) after page load. */
function selectAccount() {
  if (!msalInstance) return;
  const currentAccounts = msalInstance.getAllAccounts();
  if (currentAccounts.length === 0) {
    return;
  } else if (currentAccounts.length > 1) {
    console.warn('Multiple accounts detected, using the first one.');
    username = currentAccounts[0].username;
    homeAccountId = currentAccounts[0].homeAccountId;
    updateAuthUI(currentAccounts[0]);
  } else {
    username = currentAccounts[0].username;
    homeAccountId = currentAccounts[0].homeAccountId;
    updateAuthUI(currentAccounts[0]);
  }
}

/**
 * Processes the response from a redirect round-trip (either a login
 * or a token acquisition). Must be called once on every page load,
 * before checking for cached accounts.
 */
async function handleRedirectResponse() {
  if (!msalInstance) return;
  try {
    const response = await msalInstance.handleRedirectPromise();
    if (response !== null) {
      username = response.account.username;
      homeAccountId = response.account.homeAccountId;
      updateAuthUI(response.account);
      // If this redirect was a token request (not just a login), we
      // already have a usable Graph access token - no extra round-trip.
      if (response.accessToken) {
        pendingAccessToken = response.accessToken;
      }
    } else {
      selectAccount();
    }
  } catch (err) {
    showError('Sign-in failed: ' + err.message);
  }
}

function signIn() {
  if (!msalInstance) {
    showError(
      'No clientId set in config.js. Follow the setup steps in README.md to register a free Azure app.'
    );
    return;
  }
  msalInstance.loginRedirect(loginRequest); // navigates away; page reloads on return
}

async function signOut() {
  if (!msalInstance) return;
  const logoutRequest = {
    account: msalInstance.getAccount({ homeAccountId }),
    postLogoutRedirectUri: msalConfig.auth.redirectUri,
  };
  await msalInstance.logoutRedirect(logoutRequest); // navigates away
}

/**
 * Acquire a Graph access token: try the silent (cached) flow first,
 * fall back to a full-page redirect only when silent acquisition
 * genuinely requires user interaction.
 */
async function getTokenRedirect(request) {
  const account = msalInstance.getAccount({ homeAccountId });
  if (!account) return null;

  const fullRequest = { ...request, account };
  try {
    const result = await msalInstance.acquireTokenSilent(fullRequest);
    return result.accessToken;
  } catch (error) {
    console.warn('Silent token acquisition failed, redirecting for consent.', error);
    if (error instanceof msal.InteractionRequiredAuthError) {
      msalInstance.acquireTokenRedirect(fullRequest); // navigates away
      return null;
    }
    throw error;
  }
}


// ---- UI Helpers ---------------------------------------------

function updateAuthUI(account) {
  document.getElementById('auth-status').textContent =
    account ? (account.name || account.username) : '';
  document.getElementById('signin-btn').hidden =
    !!(account || !CONFIG.clientId);
  document.getElementById('signout-btn').hidden = !account;
}

function showLoading(visible) {
  document.getElementById('loading').hidden = !visible;
}

function showError(msg) {
  const el = document.getElementById('error-msg');
  el.innerHTML = msg;
  el.hidden = false;
}

function clearError() {
  const el = document.getElementById('error-msg');
  el.hidden = true;
  el.textContent = '';
}

// ---- Gallery ------------------------------------------------

/** Formats a Date as a group heading, e.g. "Saturday, July 12, 2025". */
function formatDateGroup(date) {
  return date.toLocaleDateString(undefined, {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
}

/** Groups items by the calendar day they were taken (falls back to
 *  "Unknown date" when no photo.takenDateTime is available), sorted
 *  chronologically (oldest day first). Within each day, photos are
 *  sorted by their exact taken time, oldest first. */
function groupItemsByDate(items) {
  const groups = new Map(); // key: 'YYYY-MM-DD' or 'unknown' -> { label, date, items }

  items.forEach((item) => {
    const taken = item.photo?.takenDateTime;
    let key = 'unknown';
    let date = null;
    let label = 'Unknown date';

    if (taken) {
      const d = new Date(taken);
      key = d.toISOString().slice(0, 10);
      date = d;
      label = formatDateGroup(d);
    }

    if (!groups.has(key)) groups.set(key, { label, date, items: [] });
    groups.get(key).items.push(item);
  });

  groups.forEach((group) => {
    group.items.sort((a, b) => {
      const ta = a.photo?.takenDateTime ? new Date(a.photo.takenDateTime) : null;
      const tb = b.photo?.takenDateTime ? new Date(b.photo.takenDateTime) : null;
      if (!ta && !tb) return 0;
      if (!ta) return 1;
      if (!tb) return -1;
      return ta - tb; // oldest first within the day
    });
  });

  return [...groups.values()].sort((a, b) => {
    if (!a.date && !b.date) return 0;
    if (!a.date) return 1; // unknown-date group last
    if (!b.date) return -1;
    return a.date - b.date; // chronological: oldest day first
  });
}

function renderGallery(items) {
  const grid = document.getElementById('photo-grid');
  const countEl = document.getElementById('photo-count');
  grid.innerHTML = '';
  countEl.textContent = `${items.length} photo${items.length !== 1 ? 's' : ''}`;

  const groups = groupItemsByDate(items);

  groups.forEach((group) => {
    const heading = document.createElement('h2');
    heading.className = 'date-heading';
    heading.textContent = `${group.label} · ${group.items.length} photo${
      group.items.length !== 1 ? 's' : ''
    }`;
    grid.appendChild(heading);

    const row = document.createElement('div');
    row.className = 'date-group-grid';

    group.items.forEach((item) => {
      const thumb = item.thumbnails?.[0]?.medium?.url ?? '';
      const hint = item.description
        ? `${item.description} (${item.name})`
        : item.name;
      const card = document.createElement('div');
      card.className = 'photo-card';
      card.setAttribute('role', 'button');
      card.setAttribute('tabindex', '0');
      card.setAttribute('aria-label', hint);
      card.title = hint;
      card.innerHTML = `<img src="${thumb}" alt="${escapeHtml(item.name)}" loading="lazy" />`;
      card.addEventListener('click', () => openLightbox(item));
      card.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') openLightbox(item);
      });
      row.appendChild(card);
    });

    grid.appendChild(row);
  });
}

function openLightbox(item) {
  const lb = document.getElementById('lightbox');
  const img = document.getElementById('lightbox-img');
  const info = document.getElementById('lightbox-info');
  const descriptionEl = document.getElementById('lightbox-description');

  // Show the thumbnail instantly so the lightbox never feels empty, then
  // swap in the full-resolution original once it's fetched (Graph doesn't
  // include @microsoft.graph.downloadUrl on the /shares children listing,
  // so it has to be requested per item, on demand).
  const placeholder =
    item.thumbnails?.[0]?.large?.url || item.thumbnails?.[0]?.medium?.url || '';
  img.src = item['@microsoft.graph.downloadUrl'] || placeholder;
  img.alt = item.name;

  if (item.description) {
    descriptionEl.textContent = item.description;
    descriptionEl.hidden = false;
  } else {
    descriptionEl.hidden = true;
  }

  const loc = item.location;
  const date = item.photo?.takenDateTime
    ? '📅 ' + new Date(item.photo.takenDateTime).toLocaleDateString()
    : '';
  const coords = loc?.latitude
    ? `📍 ${loc.latitude.toFixed(5)}, ${loc.longitude.toFixed(5)}`
    : '';

  info.innerHTML = [
    `<strong>${escapeHtml(item.name)}</strong>`,
    date,
    coords,
  ]
    .filter(Boolean)
    .join('<br>');

  lb.hidden = false;

  if (!item['@microsoft.graph.downloadUrl']) {
    fetchFullResUrl(item).then((url) => {
      // Only swap if the lightbox is still showing this same photo.
      if (url && !lb.hidden && img.alt === item.name) {
        img.src = url;
      }
    });
  }
}

/** Fetch the full-resolution download URL for a single item on demand,
 *  since Graph doesn't return @microsoft.graph.downloadUrl on the
 *  /shares/{shareId}/driveItem/children listing. */
async function fetchFullResUrl(item) {
  const driveId = item.parentReference?.driveId;
  if (!driveId || !lastAccessToken) return null;

  try {
    const url =
      `${GRAPH_API}/drives/${driveId}/items/${item.id}` +
      `?$select=@microsoft.graph.downloadUrl`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${lastAccessToken}` },
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    const fullUrl = data['@microsoft.graph.downloadUrl'];
    if (fullUrl) item['@microsoft.graph.downloadUrl'] = fullUrl; // cache on item
    return fullUrl ?? null;
  } catch {
    return null;
  }
}


function escapeHtml(str) {
  return str.replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])
  );
}

// ---- Map (Leaflet) ------------------------------------------

function initMap() {
  if (leafletMap) return;
  leafletMap = L.map('map').setView(
    CONFIG.map.defaultCenter,
    CONFIG.map.defaultZoom
  );

  const topo = L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', {
    maxZoom: 17,
    attribution:
      'Map data: © OpenStreetMap contributors, SRTM | Map style: © OpenTopoMap (CC-BY-SA)',
  });
  const streets = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '© <a href="https://openstreetmap.org">OpenStreetMap</a> contributors',
  });
  const satellite = L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    {
      maxZoom: 19,
      attribution: 'Tiles © Esri — Source: Esri, Maxar, Earthstar Geographics',
    }
  );

  topo.addTo(leafletMap);
  L.control
    .layers(
      {
        'Topographic (OpenTopoMap)': topo,
        'Streets (OSM)': streets,
        'Satellite (Esri)': satellite,
      },
      null,
      { collapsed: false }
    )
    .addTo(leafletMap);
}

function plotPhotosOnMap(items) {
  initMap();
  const geoItems = items.filter(
    (i) => i.location?.latitude != null && i.location?.longitude != null
  );
  const noGpsCount = items.length - geoItems.length;

  // Update the map legend showing how many photos lack GPS
  const legend = document.getElementById('map-legend');
  if (noGpsCount > 0) {
    legend.textContent =
      geoItems.length === 0
        ? `None of the ${items.length} photos have GPS data — switch to Gallery to browse them.`
        : `${geoItems.length} of ${items.length} photos have GPS location. ` +
          `${noGpsCount} without GPS are only visible in the Gallery.`;
    legend.hidden = false;
  } else {
    legend.hidden = true;
  }

  if (!geoItems.length) return;

  geoItems.forEach((item) => {
    const { latitude, longitude } = item.location;
    const thumbUrl = item.thumbnails?.[0]?.medium?.url ?? '';
    const icon = L.divIcon({
      className: 'map-thumb-icon',
      html: `<img src="${thumbUrl}" alt="" />`,
      iconSize: [56, 56],
      iconAnchor: [28, 28],
    });

    L.marker([latitude, longitude], { icon })
      .addTo(leafletMap)
      .on('click', () => openLightbox(item))
      .bindTooltip(item.name, { direction: 'top', offset: [0, -32] });
  });

  const bounds = L.latLngBounds(
    geoItems.map((i) => [i.location.latitude, i.location.longitude])
  );
  lastMapBounds = bounds;
  // If the map tab is hidden right now, the container has zero size and
  // fitBounds() would pick a wildly wrong zoom. Only fit immediately when
  // visible; otherwise the tab-switch handler re-applies it once shown.
  if (!document.getElementById('map-view').hidden) {
    leafletMap.fitBounds(bounds, { padding: [40, 40] });
  }
}

// ---- Main Load ----------------------------------------------

async function loadPhotos() {
  if (!CONFIG.shareUrl || CONFIG.shareUrl.includes('YOUR_SHARE_LINK')) {
    showError(
      'Please open <strong>config.js</strong> and set your OneDrive share URL. ' +
        'See <a href="README.md">README.md</a> for instructions.'
    );
    return;
  }

  clearError();

  // Microsoft Graph's /shares endpoint requires an access token even for
  // "Anyone with the link" shares, so sign-in is mandatory. Prompt rather
  // than waste a network round-trip that we know will 401.
  if (!username) {
    showError(
      'This OneDrive share requires sign-in. ' +
        (CONFIG.clientId
          ? 'Click <strong>Sign in to Microsoft</strong> above.'
          : 'Add your Azure App <code>clientId</code> to <strong>config.js</strong> first.')
    );
    document.getElementById('signin-btn').hidden = !CONFIG.clientId;
    return;
  }

  showLoading(true);

  try {
    // Reuse the token from a just-completed redirect round-trip if we
    // have one; otherwise go through the normal silent/redirect flow.
    const token = pendingAccessToken ?? (await getTokenRedirect(tokenRequest));
    pendingAccessToken = null;

    if (!token) {
      // getTokenRedirect() triggered acquireTokenRedirect(), which
      // navigates away. Nothing more to do on this page load.
      return;
    }

    lastAccessToken = token;

    const result = await fetchShareChildren(token);

    if (!result.ok) {
      showError(
        `Could not load photos (HTTP ${result.status}). ` +
          'Check that your share URL is correct and set to "Anyone with the link".'
      );
      return;
    }

    photos = result.items.filter(
      (i) => i.file && (i.file.mimeType?.startsWith('image/') || i.image)
    );

    if (!photos.length) {
      showError('No image files found in the shared folder.');
      return;
    }

    renderGallery(photos);
    plotPhotosOnMap(photos);
  } catch (err) {
    showError('Unexpected error: ' + err.message);
  } finally {
    showLoading(false);
  }
}

// ---- Event Wiring -------------------------------------------

// Keep --header-height in sync with the sticky header's real height so the
// full-window map view can size itself against it.
function updateHeaderHeightVar() {
  const header = document.getElementById('app-header');
  if (header) {
    document.documentElement.style.setProperty(
      '--header-height',
      header.offsetHeight + 'px'
    );
  }
}
window.addEventListener('resize', updateHeaderHeightVar);
updateHeaderHeightVar();

// Tabs
document.querySelectorAll('.tab').forEach((btn) => {
  btn.addEventListener('click', () => {
    const view = btn.dataset.view;
    document.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('gallery-view').hidden = view !== 'gallery';
    document.getElementById('map-view').hidden = view !== 'map';
    document.body.classList.toggle('map-mode', view === 'map');
    updateHeaderHeightVar();
    // Leaflet needs a size invalidation when its container becomes visible,
    // and bounds must be re-applied since the earlier fit may have happened
    // while the container was hidden (zero size = wrong zoom level).
    if (view === 'map' && leafletMap) {
      setTimeout(() => {
        leafletMap.invalidateSize();
        if (lastMapBounds) {
          leafletMap.fitBounds(lastMapBounds, { padding: [40, 40] });
        }
      }, 50);
    }
  });
});

// Lightbox close
document.getElementById('lightbox-close').addEventListener('click', () => {
  document.getElementById('lightbox').hidden = true;
});
document.getElementById('lightbox').addEventListener('click', (e) => {
  if (e.currentTarget === e.target) {
    document.getElementById('lightbox').hidden = true;
  }
});
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') document.getElementById('lightbox').hidden = true;
});

// Auth buttons
document.getElementById('signin-btn').addEventListener('click', signIn);
document.getElementById('signout-btn').addEventListener('click', signOut);

// ---- Boot ---------------------------------------------------
(async function init() {
  msalInstance = initMsal();
  updateAuthUI(null);
  if (msalInstance) {
    await msalInstance.initialize();
    await handleRedirectResponse(); // processes any pending redirect response
  }
  await loadPhotos();
})();
