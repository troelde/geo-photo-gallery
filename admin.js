// ============================================================
// Geo Photo Gallery — Admin App
//
// Lets the OneDrive share owner create/update/delete the
// `<photo filename>.yaml` sidecar files that app.js reads as GPS
// position / description / taken-date fallbacks for photos missing
// that metadata natively.
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

// Files.ReadWrite (not just Files.Read) so sidecar YAML files can be
// created, overwritten, and deleted.
const loginRequest = { scopes: ['User.Read', 'Files.ReadWrite'] };
const tokenRequest = { scopes: ['Files.ReadWrite'] };

let msalInstance = null;
let username = '';
let homeAccountId = null;
let pendingAccessToken = null;
let lastAccessToken = null;

let photos = []; // image items in the shared folder
let yamlByName = new Map(); // lowercased "<photo>.yaml" -> drive item
let selectedPhoto = null;
let selectedYamlItem = null; // the currently-loaded sidecar item, if any

// Centralized metadata/photos.yaml consolidation (transition phase):
// a single YAML mapping of photo filename -> {position, description,
// date}, kept in sync with the legacy per-photo sidecar files. Saves
// and deletes write/remove from both places; reads prefer this map
// when a photo has an entry here.
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
 *  (best-effort; individual sidecars still work as a fallback). */
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
 *  metadata/ folder if it doesn't exist yet, same as sidecar writes
 *  auto-create nothing (they already sit in the root). Returns true on
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

/** Fetches and parses a sidecar YAML file's content, if it exists.
 *  Returns null when there's no sidecar, it fails to fetch, or fails
 *  to parse. */
async function fetchSidecarData(yamlItem, token) {
  if (!yamlItem) return null;
  const driveId = yamlItem.parentReference?.driveId;
  if (!driveId) return null;

  try {
    const contentUrl = `${GRAPH_API}/drives/${driveId}/items/${yamlItem.id}/content`;
    const resp = await fetch(contentUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!resp.ok) return null;
    return jsyaml.load(await resp.text()) ?? {};
  } catch {
    return null;
  }
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

function yamlNameFor(photo) {
  return `${photo.name}.yaml`.toLowerCase();
}

function hasSidecar(photo) {
  return (
    yamlByName.has(yamlNameFor(photo)) ||
    yamlByName.has(`${photo.name}.yml`.toLowerCase()) ||
    centralizedKeyFor(photo) != null
  );
}

function getSidecarItem(photo) {
  return (
    yamlByName.get(yamlNameFor(photo)) ||
    yamlByName.get(`${photo.name}.yml`.toLowerCase()) ||
    null
  );
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
  filtered.forEach((photo) => {
    const li = document.createElement('li');
    li.className = 'admin-photo-item';
    if (photo === selectedPhoto) li.classList.add('active');

    const thumbUrl = photo.thumbnails?.[0]?.medium?.url ?? '';
    li.innerHTML = `
      <img src="${thumbUrl}" alt="" loading="lazy" />
      <span class="admin-photo-name">${escapeHtml(photo.name)}</span>
      ${hasSidecar(photo) ? '<span class="admin-sidecar-badge" title="Has sidecar overrides"></span>' : ''}
    `;
    li.addEventListener('click', () => selectPhoto(photo));
    list.appendChild(li);
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

function clearForm() {
  document.getElementById('admin-field-lat').value = '';
  document.getElementById('admin-field-long').value = '';
  document.getElementById('admin-field-description').value = '';
  document.getElementById('admin-field-date').value = '';
  document.getElementById('admin-field-time').value = '';
}

/** Populates the form from a parsed sidecar YAML object (position with
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
}

async function selectPhoto(photo) {
  selectedPhoto = photo;
  selectedYamlItem = getSidecarItem(photo);

  renderPhotoList(document.getElementById('admin-filter').value);

  document.getElementById('admin-detail-empty').hidden = true;
  document.getElementById('admin-detail-content').hidden = false;
  formStatus('', null);

  const thumbEl = document.getElementById('admin-detail-thumb');
  thumbEl.src = photo.thumbnails?.[0]?.medium?.url ?? '';
  document.getElementById('admin-detail-filename').textContent = photo.name;
  document.getElementById('admin-sidecar-filename').textContent = `${photo.name}.yaml`;

  const centralKey = centralizedKeyFor(photo);
  document.getElementById('admin-delete-btn').hidden = !selectedYamlItem && centralKey == null;

  renderExifInfo(photo);
  clearForm();

  if (centralKey != null) {
    // Centralized metadata is authoritative once a photo has been
    // migrated -- use it directly, no extra fetch needed.
    populateForm(centralizedData[centralKey]);
  } else if (selectedYamlItem && lastAccessToken) {
    const data = await fetchSidecarData(selectedYamlItem, lastAccessToken);
    // Guard against the user having clicked a different photo while
    // this fetch was in flight.
    if (selectedPhoto === photo) populateForm(data);
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
      'Nothing to save — fill in at least one field, or use "Remove overrides" to delete the existing sidecar.',
      'error'
    );
    return;
  }

  const saveBtn = document.getElementById('admin-save-btn');
  saveBtn.disabled = true;
  formStatus('Saving…', null);

  try {
    const driveId = selectedPhoto.parentReference?.driveId;
    const parentId = selectedPhoto.parentReference?.id;
    const yamlText = jsyaml.dump(obj);
    const url =
      `${GRAPH_API}/drives/${driveId}/items/${parentId}:/${encodeURIComponent(selectedPhoto.name)}.yaml:/content`;

    const resp = await fetch(url, {
      method: 'PUT',
      headers: {
        Authorization: 'Bearer ' + lastAccessToken,
        'Content-Type': 'text/yaml',
      },
      body: yamlText,
    });

    if (!resp.ok) {
      formStatus(`Save failed (HTTP ${resp.status}).`, 'error');
      return;
    }

    const savedItem = await resp.json();
    yamlByName.set(yamlNameFor(selectedPhoto), savedItem);
    selectedYamlItem = savedItem;

    // Transition phase: also upsert this photo's data into the
    // centralized metadata/photos.yaml file so both stay in sync.
    const centralKey = centralizedKeyFor(selectedPhoto) ?? selectedPhoto.name;
    centralizedData[centralKey] = obj;
    const centralOk = await saveCentralizedMetadata();

    document.getElementById('admin-delete-btn').hidden = false;
    renderPhotoList(document.getElementById('admin-filter').value);
    formStatus(
      centralOk
        ? 'Saved ✓ (sidecar + central metadata)'
        : 'Sidecar saved, but central metadata update failed — will retry next save.',
      centralOk ? 'success' : 'error'
    );
  } catch (err) {
    formStatus('Save failed: ' + err.message, 'error');
  } finally {
    saveBtn.disabled = false;
  }
}

async function handleDelete() {
  if (!selectedPhoto || !lastAccessToken) return;
  const centralKey = centralizedKeyFor(selectedPhoto);
  if (!selectedYamlItem && centralKey == null) return;
  if (!confirm(`Remove all sidecar overrides for "${selectedPhoto.name}"?`)) return;

  const deleteBtn = document.getElementById('admin-delete-btn');
  deleteBtn.disabled = true;
  formStatus('Removing…', null);

  try {
    if (selectedYamlItem) {
      const driveId = selectedYamlItem.parentReference?.driveId;
      const url = `${GRAPH_API}/drives/${driveId}/items/${selectedYamlItem.id}`;
      const resp = await fetch(url, {
        method: 'DELETE',
        headers: { Authorization: 'Bearer ' + lastAccessToken },
      });

      if (!resp.ok && resp.status !== 404) {
        formStatus(`Remove failed (HTTP ${resp.status}).`, 'error');
        return;
      }

      yamlByName.delete(yamlNameFor(selectedPhoto));
      selectedYamlItem = null;
    }

    let centralOk = true;
    if (centralKey != null) {
      delete centralizedData[centralKey];
      centralOk = await saveCentralizedMetadata();
    }

    clearForm();
    deleteBtn.hidden = true;
    renderPhotoList(document.getElementById('admin-filter').value);
    formStatus(
      centralOk
        ? 'Removed ✓'
        : 'Sidecar removed, but central metadata update failed — will retry next save.',
      centralOk ? 'success' : 'error'
    );
  } catch (err) {
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
      'Sign in with the OneDrive account that owns this shared folder to manage sidecar metadata. ' +
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
    yamlByName = new Map(
      result.items
        .filter((i) => i.file && /\.ya?ml$/i.test(i.name ?? ''))
        .map((i) => [i.name.toLowerCase(), i])
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
