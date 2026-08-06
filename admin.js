// ============================================================
// Geo Photo Gallery — Admin App
//
// Lets the OneDrive share owner create/update/delete per-photo entries
// in the centralized metadata/photos.yaml file that app.js reads as
// GPS position / description / taken-date fallbacks for photos
// missing that metadata natively.
//
// This is a separate, standalone page from the public gallery
// (index.html / app.js) and intentionally duplicates a small amount
// of MSAL auth boilerplate rather than refactoring the already-working
// gallery — it requests a different (write) Graph scope, so sharing
// code would add coupling without saving much.
// ============================================================

const GRAPH_API = 'https://graph.microsoft.com/v1.0';

const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: 'https://login.microsoftonline.com/consumers',
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

// Files.ReadWrite (not just Files.Read) so metadata/photos.yaml can be
// created, overwritten, and deleted.
const loginRequest = { scopes: ['User.Read', 'Files.ReadWrite'] };
const tokenRequest = { scopes: ['Files.ReadWrite'] };

let msalInstance = null;
let username = '';
let homeAccountId = null;
let pendingAccessToken = null;
let lastAccessToken = null;

let photos = []; // image items in the shared folder
let selectedPhoto = null;

// Centralized metadata/photos.yaml: a single YAML mapping of photo
// filename -> {position, description, date}. This is the sole store
// for admin-managed overrides; the legacy per-photo <name>.yaml
// sidecar files have been fully retired.
let centralizedData = {};

// ---- Share URL / Graph fetch ---------------------------------

/** Encode a OneDrive share URL into the Graph API shareId format. */
function encodeShareUrl(url) {
  const b64 = btoa(url)
    .replace(/=+$/, '')
    .replace(/\//g, '_')
    .replace(/\+/g, '-');
  return 'u!' + b64;
}

/** Same `?shareUrl=` override pattern as the public gallery. */
function getShareUrl() {
  const fromQuery = new URLSearchParams(window.location.search).get('shareUrl');
  return fromQuery || CONFIG.shareUrl;
}

/** Fetch all children (photos, yaml sidecars, metadata folder, etc.) of
 *  the shared folder, following nextLink pages. */
async function fetchFolderChildren(token) {
  const shareId = encodeShareUrl(getShareUrl());
  const select = [
    'id', 'name', 'description', 'file', 'folder', 'photo', 'location',
    'image', 'parentReference',
  ].join(',');

  const firstPage =
    `${GRAPH_API}/shares/${shareId}/driveItem/children` +
    `?$select=${select}&$expand=thumbnails($select=medium)&$top=200`;

  const headers = { Authorization: `Bearer ${token}` };
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

/** Any item's parentReference points at the shared folder's own
 *  driveId/id -- needed to path-address the metadata/ subfolder for
 *  the centralized file without a separate folder-lookup call. */
function getRootRef() {
  const anyItem = photos.find((i) => i.parentReference?.driveId) || photos[0];
  const driveId = anyItem?.parentReference?.driveId;
  const rootId = anyItem?.parentReference?.id;
  return driveId && rootId ? { driveId, rootId } : null;
}

/** Fetches the centralized metadata/photos.yaml file, if present (see
 *  the matching function in app.js -- same file/shape). Missing
 *  file/folder, a fetch failure, or malformed YAML all just yield {}
 *  (best-effort). */
async function fetchCentralizedMetadata(rootItems, token) {
  if (!token) return {};
  const anyItem = rootItems.find((i) => i.parentReference?.driveId);
  const driveId = anyItem?.parentReference?.driveId;
  const rootId = anyItem?.parentReference?.id;
  if (!driveId || !rootId) return {};

  try {
    const url =
      `${GRAPH_API}/drives/${driveId}/items/${rootId}:/metadata/photos.yaml:/content`;
    const resp = await fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
    });
    if (!resp.ok) return {};
    const data = jsyaml.load(await resp.text());
    return data && typeof data === 'object' ? data : {};
  } catch {
    return {};
  }
}

/** Writes the entire in-memory centralizedData map back to
 *  metadata/photos.yaml. Path-addressing this write auto-creates the
 *  metadata/ folder if it doesn't exist yet. Returns true on
 *  success. */
async function saveCentralizedMetadata() {
  const root = getRootRef();
  if (!root || !lastAccessToken) return false;

  try {
    const url =
      `${GRAPH_API}/drives/${root.driveId}/items/${root.rootId}:/metadata/photos.yaml:/content`;
    const resp = await fetch(url, {
      method: 'PUT',
      headers: {
        Authorization: 'Bearer ' + lastAccessToken,
        'Content-Type': 'text/yaml',
      },
      body: jsyaml.dump(centralizedData),
    });
    return resp.ok;
  } catch {
    return false;
  }
}

/** Finds this photo's key in centralizedData (case-insensitive match
 *  against the map's own keys), or null if it has no centralized
 *  entry. */
function centralizedKeyFor(photo) {
  const needle = photo.name.toLowerCase();
  return Object.keys(centralizedData).find((k) => k.toLowerCase() === needle) ?? null;
}

/** Fetches the full-resolution direct download URL for a single photo,
 *  since the shared-folder children listing doesn't reliably include
 *  @microsoft.graph.downloadUrl. Mirrors app.js's fetchFullResUrl. */
async function fetchFullResUrl(photo, token) {
  const driveId = photo.parentReference?.driveId;
  if (!driveId || !token) return null;

  try {
    const url =
      `${GRAPH_API}/drives/${driveId}/items/${photo.id}` +
      `?$select=@microsoft.graph.downloadUrl`;
    const resp = await fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    return data['@microsoft.graph.downloadUrl'] ?? null;
  } catch {
    return null;
  }
}

// ---- MSAL Auth (mirrors app.js's redirect-flow pattern) --------

function initMsal() {
  if (!CONFIG.clientId) return null;
  return new msal.PublicClientApplication(msalConfig);
}

function selectAccount() {
  if (!msalInstance) return;
  const currentAccounts = msalInstance.getAllAccounts();
  if (currentAccounts.length === 0) return;
  if (currentAccounts.length > 1) {
    console.warn('Multiple accounts detected, using the first one.');
  }
  username = currentAccounts[0].username;
  homeAccountId = currentAccounts[0].homeAccountId;
  updateAuthUI(currentAccounts[0]);
}

async function handleRedirectResponse() {
  if (!msalInstance) return;
  try {
    const response = await msalInstance.handleRedirectPromise();
    if (response !== null) {
      username = response.account.username;
      homeAccountId = response.account.homeAccountId;
      updateAuthUI(response.account);
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
  msalInstance.loginRedirect(loginRequest);
}

async function signOut() {
  if (!msalInstance) return;
  const logoutRequest = {
    account: msalInstance.getAccount({ homeAccountId }),
    postLogoutRedirectUri: msalConfig.auth.redirectUri,
  };
  await msalInstance.logoutRedirect(logoutRequest);
}

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
      msalInstance.acquireTokenRedirect(fullRequest);
      return null;
    }
    throw error;
  }
}

// ---- UI Helpers -----------------------------------------------

function updateAuthUI(account) {
  document.getElementById('admin-auth-status').textContent =
    account ? (account.name || account.username) : '';
  document.getElementById('admin-signin-btn').hidden =
    !!(account || !CONFIG.clientId);
  document.getElementById('admin-signout-btn').hidden = !account;
}

function showLoading(visible) {
  document.getElementById('admin-loading').hidden = !visible;
}

function showError(msg) {
  const el = document.getElementById('admin-error');
  el.innerHTML = msg;
  el.hidden = false;
}

function clearError() {
  const el = document.getElementById('admin-error');
  el.hidden = true;
  el.textContent = '';
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function formStatus(message, kind) {
  const el = document.getElementById('admin-form-status');
  el.textContent = message;
  el.className = 'admin-status' + (kind ? ` ${kind}` : '');
  el.hidden = !message;
}

// ---- Photo list -------------------------------------------------

function hasOverride(photo) {
  return centralizedKeyFor(photo) != null;
}

/** Same weekday/month/day formatting as the public gallery's date
 *  group headings, for visual consistency between the two apps. */
function formatDateGroup(date) {
  return date.toLocaleDateString(undefined, {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
}

/** The date used to group a photo, mirroring what the public gallery
 *  actually displays: the photo's centralized metadata/photos.yaml
 *  `date` override if it has one, else its real EXIF/OneDrive taken
 *  date, else null (grouped under "Unknown date"). Duplicates the
 *  date-string/Date-object parsing from applyYamlFallbackFields in
 *  app.js (same js-yaml quirks -- unquoted timestamps with seconds
 *  parse as Date objects, date-only/date+HH:MM stay strings). */
function effectiveTakenDate(photo) {
  const centralKey = centralizedKeyFor(photo);
  const entryDate = centralKey != null ? centralizedData[centralKey]?.date : null;

  if (entryDate != null) {
    if (entryDate instanceof Date && !isNaN(entryDate)) return entryDate;
    if (typeof entryDate === 'string') {
      const match = entryDate.trim().match(/^(\d{4}-\d{2}-\d{2})(?:[T ](\d{2}:\d{2}))?$/);
      if (match) {
        const [, datePart, timePart] = match;
        return new Date(`${datePart}T${timePart ?? '00:00'}:00Z`);
      }
    }
  }

  return photo.photo?.takenDateTime ? new Date(photo.photo.takenDateTime) : null;
}

/** Groups photos by calendar day (falls back to "Unknown date"),
 *  sorted chronologically oldest-day-first, with photos within each
 *  day sorted oldest-first -- mirroring groupItemsByDate in app.js so
 *  the admin list matches the public gallery's grouping. */
function groupPhotosByDate(photoList) {
  const groups = new Map(); // key -> { key, label, date, items }

  photoList.forEach((photo) => {
    const taken = effectiveTakenDate(photo);
    let key = 'unknown';
    let date = null;
    let label = 'Unknown date';

    if (taken && !isNaN(taken)) {
      key = taken.toISOString().slice(0, 10);
      date = taken;
      label = formatDateGroup(taken);
    }

    if (!groups.has(key)) groups.set(key, { key, label, date, items: [] });
    groups.get(key).items.push(photo);
  });

  groups.forEach((group) => {
    group.items.sort((a, b) => {
      const ta = effectiveTakenDate(a);
      const tb = effectiveTakenDate(b);
      if (!ta && !tb) return 0;
      if (!ta) return 1;
      if (!tb) return -1;
      return ta - tb;
    });
  });

  return [...groups.values()].sort((a, b) => {
    if (!a.date && !b.date) return 0;
    if (!a.date) return 1;
    if (!b.date) return -1;
    return a.date - b.date;
  });
}

function renderPhotoList(filterText = '') {
  const list = document.getElementById('admin-photo-list');
  const countEl = document.getElementById('admin-photo-count');
  const needle = filterText.trim().toLowerCase();

  const filtered = needle
    ? photos.filter((p) => p.name.toLowerCase().includes(needle))
    : photos;

  countEl.textContent = `${filtered.length} of ${photos.length} photos`;

  list.innerHTML = '';
  const groups = groupPhotosByDate(filtered);

  groups.forEach((group) => {
    const heading = document.createElement('li');
    heading.className = 'admin-date-heading';
    heading.textContent = `${group.label} · ${group.items.length} photo${
      group.items.length !== 1 ? 's' : ''
    }`;
    list.appendChild(heading);

    group.items.forEach((photo) => {
      const li = document.createElement('li');
      li.className = 'admin-photo-item';
      if (photo === selectedPhoto) li.classList.add('active');

      const thumbUrl = photo.thumbnails?.[0]?.medium?.url ?? '';
      li.innerHTML = `
        <img src="${thumbUrl}" alt="" loading="lazy" />
        <span class="admin-photo-name">${escapeHtml(photo.name)}</span>
        ${hasOverride(photo) ? '<span class="admin-sidecar-badge" title="Has metadata overrides"></span>' : ''}
      `;
      li.addEventListener('click', () => selectPhoto(photo));
      list.appendChild(li);
    });
  });
}

// ---- Detail / edit panel -----------------------------------------

function renderExifInfo(photo) {
  const dl = document.getElementById('admin-exif-fields');
  const rows = [];

  rows.push(['Description', photo.description || '—']);

  const loc = photo.location;
  rows.push([
    'GPS location',
    loc?.latitude != null
      ? `${loc.latitude.toFixed(5)}, ${loc.longitude.toFixed(5)}`
      : '—',
  ]);

  rows.push([
    'Taken date',
    photo.photo?.takenDateTime
      ? new Date(photo.photo.takenDateTime).toLocaleString(undefined, {
          dateStyle: 'medium',
          timeStyle: 'short',
        })
      : '—',
  ]);

  dl.innerHTML = rows
    .map(([label, value]) => `<dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd>`)
    .join('');
}

let positionMap = null;
let positionMarker = null;
let exifReferenceMarker = null;

const exifMarkerIcon = () =>
  L.divIcon({
    className: 'admin-exif-marker',
    html: '<span></span>',
    iconSize: [16, 16],
    iconAnchor: [8, 8],
  });

/** Lazily creates the Leaflet position-picker map the first time the
 *  detail panel becomes visible (Leaflet needs a visible container to
 *  size itself correctly). Clicking the map sets/moves the marker and
 *  fills in the Latitude/Longitude fields; dragging the marker does
 *  the same; right-click (or long-press on touch) removes the marker
 *  and clears both fields. */
function initPositionMap() {
  if (positionMap || typeof L === 'undefined') return;

  positionMap = L.map('admin-position-map', { attributionControl: false }).setView([20, 0], 2);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    subdomains: 'abc',
  }).addTo(positionMap);

  positionMap.on('click', (e) => setPositionFromLatLng(e.latlng));
  positionMap.on('contextmenu', () => {
    document.getElementById('admin-field-lat').value = '';
    document.getElementById('admin-field-long').value = '';
    updateMapMarkerFromFields();
  });
}

/** Writes a Leaflet LatLng into the Latitude/Longitude fields (rounded
 *  to 6 decimal places, ~0.1m precision) and refreshes the marker. */
function setPositionFromLatLng(latlng) {
  document.getElementById('admin-field-lat').value = latlng.lat.toFixed(6);
  document.getElementById('admin-field-long').value = latlng.lng.toFixed(6);
  updateMapMarkerFromFields();
}

/** Syncs the map marker(s) to whatever is currently in the Latitude/
 *  Longitude override fields:
 *   - When both fields hold valid numbers, shows/moves a draggable
 *     blue override marker there (and hides the EXIF reference
 *     marker, since the override takes priority).
 *   - When the fields are empty, removes the override marker and
 *     instead shows a small, non-draggable grey reference marker at
 *     the currently selected photo's real EXIF GPS location, if it
 *     has one -- this mirrors exactly what the gallery falls back to
 *     when no photos.yaml override is set for that photo.
 *  Also called after manual typing in either field so the map stays
 *  in sync either way. */
function updateMapMarkerFromFields() {
  if (!positionMap) return;

  const lat = parseFloat(document.getElementById('admin-field-lat').value);
  const lon = parseFloat(document.getElementById('admin-field-long').value);
  const hasOverride = Number.isFinite(lat) && Number.isFinite(lon);

  if (!hasOverride) {
    if (positionMarker) {
      positionMap.removeLayer(positionMarker);
      positionMarker = null;
    }

    const loc = selectedPhoto?.location;
    if (loc?.latitude != null && loc?.longitude != null) {
      if (exifReferenceMarker) {
        exifReferenceMarker.setLatLng([loc.latitude, loc.longitude]);
      } else {
        exifReferenceMarker = L.marker([loc.latitude, loc.longitude], {
          icon: exifMarkerIcon(),
          interactive: false,
        })
          .bindTooltip('EXIF location (no override set)')
          .addTo(positionMap);
      }
      positionMap.setView([loc.latitude, loc.longitude], Math.max(positionMap.getZoom(), 12));
    } else if (exifReferenceMarker) {
      positionMap.removeLayer(exifReferenceMarker);
      exifReferenceMarker = null;
    }
    return;
  }

  if (exifReferenceMarker) {
    positionMap.removeLayer(exifReferenceMarker);
    exifReferenceMarker = null;
  }

  if (positionMarker) {
    positionMarker.setLatLng([lat, lon]);
  } else {
    positionMarker = L.marker([lat, lon], { draggable: true }).addTo(positionMap);
    positionMarker.on('drag', (e) => {
      const p = e.target.getLatLng();
      document.getElementById('admin-field-lat').value = p.lat.toFixed(6);
      document.getElementById('admin-field-long').value = p.lng.toFixed(6);
    });
    positionMarker.on('contextmenu', (e) => {
      L.DomEvent.stopPropagation(e);
      document.getElementById('admin-field-lat').value = '';
      document.getElementById('admin-field-long').value = '';
      updateMapMarkerFromFields();
    });
  }
  positionMap.setView([lat, lon], Math.max(positionMap.getZoom(), 12));
}

function clearForm() {
  document.getElementById('admin-field-lat').value = '';
  document.getElementById('admin-field-long').value = '';
  document.getElementById('admin-field-description').value = '';
  document.getElementById('admin-field-date').value = '';
  document.getElementById('admin-field-time').value = '';
  updateMapMarkerFromFields();
}

/** Populates the form from a parsed metadata YAML object (position with
 *  lat/long or latitude/longitude aliases; description; date as
 *  YYYY-MM-DD or YYYY-MM-DDTHH:MM, or a Date object if js-yaml parsed a
 *  timestamp-with-seconds form). */
function populateForm(data) {
  clearForm();
  if (!data) return;

  const pos = data.position ?? {};
  const lat = pos.lat ?? pos.latitude;
  const lon = pos.long ?? pos.lon ?? pos.lng ?? pos.longitude;
  if (typeof lat === 'number') document.getElementById('admin-field-lat').value = lat;
  if (typeof lon === 'number') document.getElementById('admin-field-long').value = lon;

  if (typeof data.description === 'string') {
    document.getElementById('admin-field-description').value = data.description;
  }

  let datePart = '';
  let timePart = '';
  if (data.date instanceof Date && !isNaN(data.date)) {
    const iso = data.date.toISOString();
    datePart = iso.slice(0, 10);
    timePart = iso.slice(11, 16);
  } else if (typeof data.date === 'string') {
    const match = data.date.trim().match(/^(\d{4}-\d{2}-\d{2})(?:[T ](\d{2}:\d{2}))?/);
    if (match) {
      datePart = match[1];
      timePart = match[2] ?? '';
    }
  }
  document.getElementById('admin-field-date').value = datePart;
  document.getElementById('admin-field-time').value = timePart;
  updateMapMarkerFromFields();
}

function selectPhoto(photo) {
  selectedPhoto = photo;

  renderPhotoList(document.getElementById('admin-filter').value);

  document.getElementById('admin-detail-empty').hidden = true;
  document.getElementById('admin-detail-content').hidden = false;
  formStatus('', null);

  const thumbEl = document.getElementById('admin-detail-thumb');
  thumbEl.src = photo.thumbnails?.[0]?.medium?.url ?? '';
  document.getElementById('admin-detail-filename').textContent = photo.name;

  const centralKey = centralizedKeyFor(photo);
  document.getElementById('admin-delete-btn').hidden = centralKey == null;

  renderExifInfo(photo);

  // The map container was hidden (inside #admin-detail-content) until
  // now, so Leaflet needs an explicit init/resize once it's visible.
  // Must happen before populateForm() below so its marker/centering
  // update (which reads selectedPhoto.location for the EXIF fallback
  // pin) has a live map to draw onto.
  initPositionMap();
  if (positionMap) positionMap.invalidateSize();

  populateForm(centralKey != null ? centralizedData[centralKey] : null);

  if (positionMap && !positionMarker && !exifReferenceMarker) {
    // Neither an override nor an EXIF location exists for this photo.
    positionMap.setView([20, 0], 2);
  }

  // Swap in the full-resolution original once fetched, so the preview
  // isn't limited to OneDrive's small "medium" thumbnail size.
  if (lastAccessToken) {
    fetchFullResUrl(photo, lastAccessToken).then((url) => {
      if (url && selectedPhoto === photo) thumbEl.src = url;
    });
  }
}

// ---- Save / Delete ------------------------------------------------

/** Builds a plain object from only the populated form fields, matching
 *  the position.lat/long + description + date convention app.js reads. */
function buildSidecarObject() {
  const lat = parseFloat(document.getElementById('admin-field-lat').value);
  const lon = parseFloat(document.getElementById('admin-field-long').value);
  const description = document.getElementById('admin-field-description').value.trim();
  const date = document.getElementById('admin-field-date').value; // '' or YYYY-MM-DD
  const time = document.getElementById('admin-field-time').value; // '' or HH:MM

  const obj = {};
  if (Number.isFinite(lat) && Number.isFinite(lon)) {
    obj.position = { lat, long: lon };
  }
  if (description) {
    obj.description = description;
  }
  if (date) {
    obj.date = time ? `${date}T${time}` : date;
  }
  return obj;
}

async function handleSave(e) {
  e.preventDefault();
  if (!selectedPhoto || !lastAccessToken) return;

  const obj = buildSidecarObject();
  if (Object.keys(obj).length === 0) {
    formStatus(
      'Nothing to save — fill in at least one field, or use "Remove overrides" to delete the existing override.',
      'error'
    );
    return;
  }

  const saveBtn = document.getElementById('admin-save-btn');
  saveBtn.disabled = true;
  formStatus('Saving…', null);

  try {
    const centralKey = centralizedKeyFor(selectedPhoto) ?? selectedPhoto.name;
    centralizedData[centralKey] = obj;
    const ok = await saveCentralizedMetadata();

    if (!ok) {
      delete centralizedData[centralKey];
      formStatus('Save failed — could not write metadata/photos.yaml.', 'error');
      return;
    }

    document.getElementById('admin-delete-btn').hidden = false;
    renderPhotoList(document.getElementById('admin-filter').value);
    formStatus('Saved ✓', 'success');
  } catch (err) {
    formStatus('Save failed: ' + err.message, 'error');
  } finally {
    saveBtn.disabled = false;
  }
}

async function handleDelete() {
  if (!selectedPhoto || !lastAccessToken) return;
  const centralKey = centralizedKeyFor(selectedPhoto);
  if (centralKey == null) return;
  if (!confirm(`Remove all metadata overrides for "${selectedPhoto.name}"?`)) return;

  const deleteBtn = document.getElementById('admin-delete-btn');
  deleteBtn.disabled = true;
  formStatus('Removing…', null);

  const removedEntry = centralizedData[centralKey];
  try {
    delete centralizedData[centralKey];
    const ok = await saveCentralizedMetadata();

    if (!ok) {
      centralizedData[centralKey] = removedEntry;
      formStatus('Remove failed — could not write metadata/photos.yaml.', 'error');
      return;
    }

    clearForm();
    deleteBtn.hidden = true;
    renderPhotoList(document.getElementById('admin-filter').value);
    formStatus('Removed ✓', 'success');
  } catch (err) {
    centralizedData[centralKey] = removedEntry;
    formStatus('Remove failed: ' + err.message, 'error');
  } finally {
    deleteBtn.disabled = false;
  }
}

// ---- Main load ----------------------------------------------------

async function loadAdminData() {
  const shareUrl = getShareUrl();
  if (!shareUrl || shareUrl.includes('YOUR_SHARE_LINK')) {
    showError(
      'Please open <strong>config.js</strong> and set your OneDrive share URL ' +
        '(or pass one via <code>?shareUrl=</code> in the page URL).'
    );
    return;
  }

  clearError();

  if (!username) {
    showError(
      'Sign in with the OneDrive account that owns this shared folder to manage metadata overrides. ' +
        (CONFIG.clientId
          ? 'Click <strong>Sign in to Microsoft</strong> above.'
          : 'Add your Azure App <code>clientId</code> to <strong>config.js</strong> first.')
    );
    document.getElementById('admin-signin-btn').hidden = !CONFIG.clientId;
    return;
  }

  showLoading(true);
  document.getElementById('admin-main').hidden = true;

  try {
    const token = pendingAccessToken || (await getTokenRedirect(tokenRequest));
    if (!token) return; // acquireTokenRedirect() is navigating away
    lastAccessToken = token;

    const result = await fetchFolderChildren(token);
    if (!result.ok) {
      showError(
        `Could not load photos (HTTP ${result.status}). ` +
          'Check that your share URL is correct and that this account has ' +
          'edit access to the shared folder.'
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

    centralizedData = await fetchCentralizedMetadata(result.items, token);

    document.getElementById('admin-main').hidden = false;
    renderPhotoList();
  } catch (err) {
    showError('Unexpected error: ' + err.message);
  } finally {
    showLoading(false);
  }
}

// ---- Event wiring ---------------------------------------------------

document.getElementById('admin-signin-btn').addEventListener('click', signIn);
document.getElementById('admin-signout-btn').addEventListener('click', signOut);
document.getElementById('admin-sidecar-form').addEventListener('submit', handleSave);
document.getElementById('admin-delete-btn').addEventListener('click', handleDelete);
document.getElementById('admin-filter').addEventListener('input', (e) => {
  renderPhotoList(e.target.value);
});
document.getElementById('admin-field-lat').addEventListener('input', updateMapMarkerFromFields);
document.getElementById('admin-field-long').addEventListener('input', updateMapMarkerFromFields);

// ---- Boot -------------------------------------------------------------

(async function init() {
  msalInstance = initMsal();
  updateAuthUI(null);
  if (msalInstance) {
    await msalInstance.initialize();
    await handleRedirectResponse();
  }
  await loadAdminData();
})();
