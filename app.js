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
let photoMarkers = new Map(); // item -> Leaflet marker, so the lightbox can jump to a photo's pin
let photos = [];
let sortedPhotos = []; // chronologically flattened order, used for lightbox prev/next
let currentPhotoIndex = -1; // index into sortedPhotos of the currently open lightbox photo
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

/** Resolves the OneDrive share URL to use: a `?shareUrl=` query parameter
 *  takes priority (lets you link to a different shared folder without
 *  editing config.js), falling back to CONFIG.shareUrl otherwise. */
function getShareUrl() {
  const fromQuery = new URLSearchParams(window.location.search).get('shareUrl');
  return fromQuery || CONFIG.shareUrl;
}


/** Fetch all image children of a shared folder, following nextLink pages. */
async function fetchShareChildren(token) {
  const shareId = encodeShareUrl(getShareUrl());
  const select = [
    'id', 'name', 'description', 'file', 'folder', 'photo', 'location',
    'image', 'parentReference', '@microsoft.graph.downloadUrl',
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

/** Looks inside `metadata/` for:
 *   - `description.md` — an overall gallery intro/description
 *   - `YYYYMMDD-description.md` — a per-day description, shown under that
 *     day's heading in the gallery (YYYYMMDD must match a date group's key)
 *  Returns { global: string|null, byDate: Map<'YYYY-MM-DD', string> }.
 *  Missing folder/files or any failure just yields empty results — this
 *  is a nice-to-have, not required for the gallery to function. */
async function fetchDescriptions(rootItems, token) {
  const empty = { global: null, byDate: new Map() };

  const metadataFolder = rootItems.find(
    (i) => i.folder && i.name?.toLowerCase() === 'metadata'
  );
  const driveId = metadataFolder?.parentReference?.driveId;
  if (!metadataFolder || !driveId || !token) return empty;

  try {
    const childrenUrl =
      `${GRAPH_API}/drives/${driveId}/items/${metadataFolder.id}/children` +
      `?$select=id,name,file&$top=200`;
    const childrenResp = await fetch(childrenUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!childrenResp.ok) return empty;
    const childrenData = await childrenResp.json();
    const files = (childrenData.value ?? []).filter((c) => c.file);

    const fetchText = async (itemId) => {
      const contentUrl = `${GRAPH_API}/drives/${driveId}/items/${itemId}/content`;
      const resp = await fetch(contentUrl, {
        headers: { Authorization: `Bearer ${token}` },
      });
      return resp.ok ? await resp.text() : null;
    };

    const globalFile = files.find(
      (c) => c.name?.toLowerCase() === 'description.md'
    );
    const dateFileRegex = /^(\d{8})-description\.md$/i;
    const dateFiles = files
      .map((c) => ({ file: c, match: c.name?.match(dateFileRegex) }))
      .filter((x) => x.match);

    const [globalText, ...dateTexts] = await Promise.all([
      globalFile ? fetchText(globalFile.id) : Promise.resolve(null),
      ...dateFiles.map((x) => fetchText(x.file.id)),
    ]);

    const byDate = new Map();
    dateFiles.forEach((x, i) => {
      const text = dateTexts[i];
      if (!text) return;
      const raw = x.match[1]; // YYYYMMDD
      const key = `${raw.slice(0, 4)}-${raw.slice(4, 6)}-${raw.slice(6, 8)}`;
      byDate.set(key, text);
    });

    return { global: globalText, byDate };
  } catch {
    return empty;
  }
}

/** Applies GPS/description/date override fields from a parsed YAML
 *  data object (the centralized metadata/photos.yaml entry) onto a
 *  photo, in place. Any field present (non-null) in the YAML data
 *  overrides the photo's real EXIF/OneDrive value for that field --
 *  this lets photos.yaml both fill gaps (missing EXIF) and correct
 *  wrong/undesired EXIF data. Fields omitted from the YAML entry are
 *  left untouched. */
function applyYamlFallbackFields(photo, data) {
  if (!data) return;

  const pos = data.position ?? {};
  const lat = pos.lat ?? pos.latitude;
  const lon = pos.long ?? pos.lon ?? pos.lng ?? pos.longitude;
  if (typeof lat === 'number' && typeof lon === 'number') {
    photo.location = { latitude: lat, longitude: lon };
  }

  if (typeof data.description === 'string') {
    photo.description = data.description;
  }

  if (data.date != null) {
    let isoDateTime = null;
    if (data.date instanceof Date && !isNaN(data.date)) {
      // js-yaml parses unquoted timestamps that include seconds
      // (e.g. YYYY-MM-DDTHH:MM:SS) as Date objects rather than
      // strings.
      isoDateTime = data.date.toISOString();
    } else if (typeof data.date === 'string') {
      // Accepts "YYYY-MM-DD" (date only) or "YYYY-MM-DDTHH:MM"
      // (date + time, no seconds) -- both stay strings under
      // js-yaml since neither matches its full timestamp regex.
      const match = data.date
        .trim()
        .match(/^(\d{4}-\d{2}-\d{2})(?:[T ](\d{2}:\d{2}))?$/);
      if (match) {
        const [, datePart, timePart] = match;
        isoDateTime = `${datePart}T${timePart ?? '00:00'}:00Z`;
      }
    }
    if (isoDateTime) {
      photo.photo = { ...photo.photo, takenDateTime: isoDateTime };
    }
  }
}

/** Fetches the centralized metadata/photos.yaml file, if present: a
 *  single YAML mapping of photo filename -> {position, description,
 *  date}. This is the sole source of GPS/description/date fallback
 *  data for the gallery (legacy per-photo <name>.yaml sidecars are no
 *  longer read here -- see applyYamlFallbacks below). Path-addressed
 *  directly off the shared folder's own root item, so no folder
 *  search is needed -- any child's parentReference already points at
 *  that root driveId/id. Missing file/folder, a fetch failure, or
 *  malformed YAML all just yield {} (best-effort). */
async function fetchCentralizedMetadata(rootItems, token) {
  if (!token || typeof jsyaml === 'undefined') return {};

  const anyItem = rootItems.find((i) => i.parentReference?.driveId);
  const driveId = anyItem?.parentReference?.driveId;
  const rootId = anyItem?.parentReference?.id;
  if (!driveId || !rootId) return {};

  try {
    const url =
      `${GRAPH_API}/drives/${driveId}/items/${rootId}:/metadata/photos.yaml:/content`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!resp.ok) return {};
    const data = jsyaml.load(await resp.text());
    return data && typeof data === 'object' ? data : {};
  } catch {
    return {};
  }
}

/** Looks up per-photo overrides in the centralized metadata/photos.yaml
 *  file (see fetchCentralizedMetadata) and applies them onto matching
 *  photos, in place. Per-photo entries can provide:
 *   - a `position` dict with `lat`/`long` (or `latitude`/`longitude`,
 *     decimal degrees) used as the photo's GPS location so it shows up
 *     on the map;
 *   - a `description` string used as the photo's subtitle;
 *   - a `date` string (`YYYY-MM-DD` or `YYYY-MM-DDTHH:MM`) used as the
 *     photo's taken date/time (for date-grouping/sorting).
 *  Any of these present (non-null) in a photo's YAML entry overrides
 *  that photo's real OneDrive/EXIF value for the same field -- this
 *  both fills gaps (missing EXIF) and lets you correct wrong/undesired
 *  EXIF data. Fields omitted from the YAML entry are left untouched. */
function applyYamlFallbacks(photos, centralizedData) {
  const centralizedByName = new Map(
    Object.keys(centralizedData || {}).map((name) => [
      name.toLowerCase(),
      centralizedData[name],
    ])
  );
  if (!centralizedByName.size) return;

  for (const photo of photos) {
    const centralEntry = centralizedByName.get(photo.name.toLowerCase());
    if (centralEntry) applyYamlFallbackFields(photo, centralEntry);
  }
}

/** Renders fetched markdown (if any) into the gallery description box. */
function renderGalleryDescription(markdown) {
  const el = document.getElementById('gallery-description');
  if (!markdown || typeof marked === 'undefined') {
    el.hidden = true;
    return;
  }
  el.innerHTML = marked.parse(markdown, { renderer: markdownLinkRenderer() });
  el.hidden = false;
}

/** A marked.js renderer that makes every link open in a new tab, so
 *  clicking a link in the gallery description doesn't navigate away
 *  from the gallery itself. */
function markdownLinkRenderer() {
  const renderer = new marked.Renderer();
  const defaultLink = renderer.link.bind(renderer);
  renderer.link = (href, title, text) =>
    defaultLink(href, title, text).replace(
      '<a ',
      '<a target="_blank" rel="noopener noreferrer" '
    );
  return renderer;
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

    if (!groups.has(key)) groups.set(key, { key, label, date, items: [] });
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

function renderGallery(items, dateDescriptions = new Map()) {
  const grid = document.getElementById('photo-grid');
  const countEl = document.getElementById('photo-count');
  grid.innerHTML = '';
  countEl.textContent = `${items.length} photo${items.length !== 1 ? 's' : ''}`;

  const groups = groupItemsByDate(items);
  sortedPhotos = groups.flatMap((g) => g.items);

  groups.forEach((group) => {
    const heading = document.createElement('h2');
    heading.className = 'date-heading';
    heading.textContent = `${group.label} · ${group.items.length} photo${
      group.items.length !== 1 ? 's' : ''
    }`;
    grid.appendChild(heading);

    const dayMarkdown = dateDescriptions.get(group.key);
    if (dayMarkdown && typeof marked !== 'undefined') {
      const dayDesc = document.createElement('div');
      dayDesc.className = 'date-description markdown-body';
      dayDesc.innerHTML = marked.parse(dayMarkdown, {
        renderer: markdownLinkRenderer(),
      });
      grid.appendChild(dayDesc);
    }

    const row = document.createElement('div');
    row.className = 'date-group-grid';

    group.items.forEach((item) => {
      const thumb = item.thumbnails?.[0]?.medium?.url ?? '';
      const hint = item.description || item.name;
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
  currentPhotoIndex = sortedPhotos.indexOf(item);
  showLightboxItem(item);
  document.getElementById('lightbox').hidden = false;
  const multiplePhotos = sortedPhotos.length > 1;
  document.getElementById('lightbox-prev').disabled = !multiplePhotos;
  document.getElementById('lightbox-next').disabled = !multiplePhotos;
}

function showNextPhoto() {
  if (!sortedPhotos.length) return;
  currentPhotoIndex = (currentPhotoIndex + 1) % sortedPhotos.length;
  showLightboxItem(sortedPhotos[currentPhotoIndex]);
}

function showPrevPhoto() {
  if (!sortedPhotos.length) return;
  currentPhotoIndex =
    (currentPhotoIndex - 1 + sortedPhotos.length) % sortedPhotos.length;
  showLightboxItem(sortedPhotos[currentPhotoIndex]);
}

function showLightboxItem(item) {
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
    ? '📅 ' +
      new Date(item.photo.takenDateTime).toLocaleString(undefined, {
        dateStyle: 'medium',
        timeStyle: 'short',
      })
    : '';

  info.innerHTML = [`<strong>${escapeHtml(item.name)}</strong>`, date]
    .filter(Boolean)
    .join('<br>');

  const viewOnMapBtn = document.getElementById('lightbox-view-on-map');
  const hasGps = loc?.latitude != null && loc?.longitude != null;
  viewOnMapBtn.hidden = !hasGps;
  viewOnMapBtn.onclick = hasGps ? () => showPhotoOnMap(item) : null;
  viewOnMapBtn.title = hasGps
    ? `${loc.latitude.toFixed(5)}, ${loc.longitude.toFixed(5)}`
    : '';

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

  photoMarkers.clear();

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

    const marker = L.marker([latitude, longitude], { icon })
      .addTo(leafletMap)
      .on('click', () => openLightbox(item))
      .bindTooltip(item.name, { direction: 'top', offset: [0, -32] });

    photoMarkers.set(item, marker);
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
  const shareUrl = getShareUrl();
  if (!shareUrl || shareUrl.includes('YOUR_SHARE_LINK')) {
    showError(
      'Please open <strong>config.js</strong> and set your OneDrive share URL ' +
        '(or pass one via <code>?shareUrl=</code> in the page URL). ' +
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

    // Look for optional metadata/description.md (gallery-wide) and
    // metadata/YYYYMMDD-description.md (per-day) files, and the
    // centralized metadata/photos.yaml file providing per-photo
    // GPS/description/date overrides, before rendering.
    const [descriptions, centralizedData] = await Promise.all([
      fetchDescriptions(result.items, token),
      fetchCentralizedMetadata(result.items, token),
    ]);
    applyYamlFallbacks(photos, centralizedData);

    renderGallery(photos, descriptions.byDate);
    plotPhotosOnMap(photos);
    renderGalleryDescription(descriptions.global);
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
function switchView(view) {
  document.querySelectorAll('.tab').forEach((t) =>
    t.classList.toggle('active', t.dataset.view === view)
  );
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
}

document.querySelectorAll('.tab').forEach((btn) => {
  btn.addEventListener('click', () => switchView(btn.dataset.view));
});

/** Switch to the Map tab, center on the given photo's marker, and open
 *  its tooltip so it's obvious which pin corresponds to the photo that
 *  was just being viewed in the lightbox. */
function showPhotoOnMap(item) {
  const marker = photoMarkers.get(item);
  if (!marker) return;
  document.getElementById('lightbox').hidden = true;
  switchView('map');
  setTimeout(() => {
    leafletMap.setView(marker.getLatLng(), Math.max(leafletMap.getZoom(), 15));
    marker.openTooltip();
  }, 60);
}

// Lightbox close
document.getElementById('lightbox-close').addEventListener('click', () => {
  document.getElementById('lightbox').hidden = true;
});
document.getElementById('lightbox').addEventListener('click', (e) => {
  if (e.currentTarget === e.target) {
    document.getElementById('lightbox').hidden = true;
  }
});
document.getElementById('lightbox-prev').addEventListener('click', (e) => {
  e.stopPropagation();
  showPrevPhoto();
});
document.getElementById('lightbox-next').addEventListener('click', (e) => {
  e.stopPropagation();
  showNextPhoto();
});
document.addEventListener('keydown', (e) => {
  if (document.getElementById('lightbox').hidden) return;
  if (e.key === 'Escape') document.getElementById('lightbox').hidden = true;
  if (e.key === 'ArrowRight') showNextPhoto();
  if (e.key === 'ArrowLeft') showPrevPhoto();
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
