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

// The overall gallery description shown at the top of the public
// gallery (metadata/description.md), and the text last successfully
// saved -- used to detect unsaved changes.
let galleryDescriptionText = '';
let lastSavedDescriptionText = '';

// Per-day descriptions (metadata/YYYYMMDD-description.md), one per
// calendar day that has at least one photo. dayDescriptionDates is the
// list of candidate days derived from the photos themselves;
// dayDescriptionTexts maps 'YYYY-MM-DD' -> saved text (only present if
// the file already exists in OneDrive -- missing entries mean "not
// created yet", and Save will create the file).
let dayDescriptionDates = []; // [{ key, label, photoCount }]
let dayDescriptionTexts = {}; // 'YYYY-MM-DD' -> existing text
let selectedDayKey = null;
let lastSavedDayText = '';

// Which top-level tab ("photos", "description", or "daydesc") is
// currently shown.
let activeTab = 'photos';

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
    `?$select=${select}&$expand=thumbnails($select=medium,large)&$top=200`;

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

/** Fetches the overall gallery description (see the matching function
 *  in app.js -- same file/shape: plain Markdown text). Missing
 *  file/folder or a fetch failure just yields '' (best-effort). */
async function fetchGalleryDescription(rootItems, token) {
  if (!token) return '';
  const anyItem = rootItems.find((i) => i.parentReference?.driveId);
  const driveId = anyItem?.parentReference?.driveId;
  const rootId = anyItem?.parentReference?.id;
  if (!driveId || !rootId) return '';

  try {
    const url =
      `${GRAPH_API}/drives/${driveId}/items/${rootId}:/metadata/description.md:/content`;
    const resp = await fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
    });
    if (!resp.ok) return '';
    return await resp.text();
  } catch {
    return '';
  }
}

/** Writes the given Markdown text back to metadata/description.md.
 *  Path-addressing this write auto-creates the metadata/ folder if it
 *  doesn't exist yet. Returns true on success. */
async function saveGalleryDescription(text) {
  const root = getRootRef();
  if (!root || !lastAccessToken) return false;

  try {
    const url =
      `${GRAPH_API}/drives/${root.driveId}/items/${root.rootId}:/metadata/description.md:/content`;
    const resp = await fetch(url, {
      method: 'PUT',
      headers: {
        Authorization: 'Bearer ' + lastAccessToken,
        'Content-Type': 'text/markdown',
      },
      body: text,
    });
    return resp.ok;
  } catch {
    return false;
  }
}

/** Converts a 'YYYY-MM-DD' date key to the compact YYYYMMDD form used
 *  in per-day description filenames (see the matching regex in
 *  app.js's fetchDescriptions). */
function compactDateKey(dateKey) {
  return dateKey.replace(/-/g, '');
}

/** Fetches every existing metadata/YYYYMMDD-description.md file's
 *  content (see the matching function in app.js -- same file/shape:
 *  plain Markdown text per calendar day). Returns a
 *  Map<'YYYY-MM-DD', string>; days with no file simply have no entry
 *  (missing folder/files or any failure just yields an empty map --
 *  best-effort, same as app.js). */
async function fetchDayDescriptions(rootItems, token) {
  const empty = new Map();
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
      headers: { Authorization: 'Bearer ' + token },
    });
    if (!childrenResp.ok) return empty;
    const childrenData = await childrenResp.json();
    const files = (childrenData.value ?? []).filter((c) => c.file);

    const dateFileRegex = /^(\d{8})-description\.md$/i;
    const matches = files
      .map((c) => ({ file: c, match: c.name?.match(dateFileRegex) }))
      .filter((x) => x.match);

    const texts = await Promise.all(
      matches.map(async (x) => {
        const contentUrl = `${GRAPH_API}/drives/${driveId}/items/${x.file.id}/content`;
        const resp = await fetch(contentUrl, {
          headers: { Authorization: 'Bearer ' + token },
        });
        return resp.ok ? await resp.text() : '';
      })
    );

    const map = new Map();
    matches.forEach((x, i) => {
      const raw = x.match[1]; // YYYYMMDD
      const key = `${raw.slice(0, 4)}-${raw.slice(4, 6)}-${raw.slice(6, 8)}`;
      map.set(key, texts[i]);
    });
    return map;
  } catch {
    return empty;
  }
}

/** Writes the given Markdown text to metadata/YYYYMMDD-description.md
 *  for the given 'YYYY-MM-DD' date key. Path-addressing this write
 *  auto-creates the metadata/ folder and/or the file itself if either
 *  doesn't exist yet -- so this both creates missing per-day files and
 *  updates existing ones. Returns true on success. */
async function saveDayDescription(dateKey, text) {
  const root = getRootRef();
  if (!root || !lastAccessToken) return false;

  try {
    const filename = `${compactDateKey(dateKey)}-description.md`;
    const url =
      `${GRAPH_API}/drives/${root.driveId}/items/${root.rootId}:/metadata/${filename}:/content`;
    const resp = await fetch(url, {
      method: 'PUT',
      headers: {
        Authorization: 'Bearer ' + lastAccessToken,
        'Content-Type': 'text/markdown',
      },
      body: text,
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

/** Switches between the "photos" and "description" top-level tabs.
 *  Both panels are always fully loaded together (see loadAdminData);
 *  this only toggles which one is visible and highlights the active
 *  tab button. */
function switchTab(tab) {
  activeTab = tab;
  document.getElementById('admin-main').hidden = tab !== 'photos';
  document.getElementById('admin-description-section').hidden = tab !== 'description';
  document.getElementById('admin-daydesc-section').hidden = tab !== 'daydesc';
  document.getElementById('admin-pdf-section').hidden = tab !== 'pdf';
  document.querySelectorAll('.admin-tab-btn').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.tab === tab);
  });
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

function descriptionStatus(message, kind) {
  const el = document.getElementById('admin-description-status');
  el.textContent = message;
  el.className = 'admin-status' + (kind ? ` ${kind}` : '');
  el.hidden = !message;
}

// ---- Gallery description editor ---------------------------------

/** A marked.js renderer that makes every link open in a new tab,
 *  matching the public gallery's own rendering (see app.js). */
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

/** Re-renders the live Markdown preview from the current textarea
 *  contents. Empty text just clears the preview (CSS shows a hint). */
function renderDescriptionPreview() {
  const preview = document.getElementById('admin-description-preview');
  const text = document.getElementById('admin-description-text').value;
  if (!text.trim() || typeof marked === 'undefined') {
    preview.innerHTML = '';
    return;
  }
  preview.innerHTML = marked.parse(text, { renderer: markdownLinkRenderer() });
}

async function handleSaveDescription() {
  const text = document.getElementById('admin-description-text').value;
  descriptionStatus('Saving…', '');
  const saveBtn = document.getElementById('admin-description-save-btn');
  saveBtn.disabled = true;
  try {
    const ok = await saveGalleryDescription(text);
    if (ok) {
      galleryDescriptionText = text;
      lastSavedDescriptionText = text;
      descriptionStatus('Saved ✓', 'success');
    } else {
      descriptionStatus(
        'Save failed — could not write metadata/description.md.',
        'error'
      );
    }
  } catch (err) {
    descriptionStatus('Save failed: ' + err.message, 'error');
  } finally {
    saveBtn.disabled = false;
  }
}

// ---- Per-day description editor ----------------------------------

function dayDescStatus(message, kind) {
  const el = document.getElementById('admin-daydesc-status');
  el.textContent = message;
  el.className = 'admin-status' + (kind ? ` ${kind}` : '');
  el.hidden = !message;
}

/** Derives the list of candidate per-day description files: one entry
 *  per distinct calendar day among the loaded photos (using the same
 *  effective-date/grouping logic as the photo list), oldest day
 *  first. Photos with no determinable date ("Unknown date") are
 *  excluded -- there's no single YYYYMMDD to file them under. */
function computeDayDescriptionDates() {
  return groupPhotosByDate(photos)
    .filter((group) => group.date)
    .map((group) => ({
      key: group.key,
      label: group.label,
      photoCount: group.items.length,
    }));
}

/** Re-renders the live Markdown preview for the currently selected
 *  day's textarea contents. */
function renderDayDescPreview() {
  const preview = document.getElementById('admin-daydesc-preview');
  const text = document.getElementById('admin-daydesc-text').value;
  if (!text.trim() || typeof marked === 'undefined') {
    preview.innerHTML = '';
    return;
  }
  preview.innerHTML = marked.parse(text, { renderer: markdownLinkRenderer() });
}

function renderDayDescriptionList() {
  const list = document.getElementById('admin-daydesc-list');
  const countEl = document.getElementById('admin-daydesc-count');
  countEl.textContent = `${dayDescriptionDates.length} day${
    dayDescriptionDates.length !== 1 ? 's' : ''
  } with photos`;

  list.innerHTML = '';
  dayDescriptionDates.forEach((day) => {
    const exists = Object.prototype.hasOwnProperty.call(dayDescriptionTexts, day.key);
    const li = document.createElement('li');
    li.className = 'admin-daydesc-item';
    if (day.key === selectedDayKey) li.classList.add('active');
    li.innerHTML = `
      <span class="admin-daydesc-item-label">${escapeHtml(day.label)}</span>
      <span class="admin-daydesc-item-count">${day.photoCount} photo${day.photoCount !== 1 ? 's' : ''}</span>
      ${exists ? '<span class="admin-daydesc-exists-badge" title="Description file exists"></span>' : ''}
    `;
    li.addEventListener('click', () => selectDayDescription(day));
    list.appendChild(li);
  });
}

/** Selects a day in the list and populates the editor with its
 *  existing text (or empty, if the file doesn't exist yet -- Save
 *  will then create it). */
function selectDayDescription(day) {
  selectedDayKey = day.key;
  const exists = Object.prototype.hasOwnProperty.call(dayDescriptionTexts, day.key);
  const text = exists ? dayDescriptionTexts[day.key] : '';
  lastSavedDayText = text;

  document.getElementById('admin-daydesc-heading').textContent = day.label;
  document.getElementById('admin-daydesc-file-status').textContent = exists
    ? `Editing existing metadata/${compactDateKey(day.key)}-description.md`
    : `No file yet — saving will create metadata/${compactDateKey(day.key)}-description.md`;
  document.getElementById('admin-daydesc-text').value = text;
  renderDayDescPreview();
  dayDescStatus('');

  document.getElementById('admin-daydesc-empty').hidden = true;
  document.getElementById('admin-daydesc-content').hidden = false;

  renderDayDescriptionList();
}

async function handleSaveDayDescription() {
  if (!selectedDayKey) return;
  const text = document.getElementById('admin-daydesc-text').value;
  const wasNew = !Object.prototype.hasOwnProperty.call(dayDescriptionTexts, selectedDayKey);
  dayDescStatus('Saving…', '');
  const saveBtn = document.getElementById('admin-daydesc-save-btn');
  saveBtn.disabled = true;
  try {
    const ok = await saveDayDescription(selectedDayKey, text);
    if (ok) {
      dayDescriptionTexts[selectedDayKey] = text;
      lastSavedDayText = text;
      document.getElementById('admin-daydesc-file-status').textContent =
        `Editing existing metadata/${compactDateKey(selectedDayKey)}-description.md`;
      dayDescStatus(wasNew ? 'Created ✓' : 'Saved ✓', 'success');
      renderDayDescriptionList();
    } else {
      dayDescStatus(
        `Save failed — could not write metadata/${compactDateKey(selectedDayKey)}-description.md.`,
        'error'
      );
    }
  } catch (err) {
    dayDescStatus('Save failed: ' + err.message, 'error');
  } finally {
    saveBtn.disabled = false;
  }
}

// ---- PDF export ------------------------------------------------

/** The effective description shown for a photo: the centralized
 *  metadata/photos.yaml override if present and non-blank, else the
 *  photo's real OneDrive/EXIF description, else null. A blank/
 *  whitespace-only override does NOT count as an override (falls
 *  back to the real description instead) -- this avoids silently
 *  hiding real captions for photos whose photos.yaml entry happens to
 *  have an empty `description:` key. Mirrors the override precedence
 *  in applyYamlFallbackFields in app.js. */
function effectiveDescription(photo) {
  const centralKey = centralizedKeyFor(photo);
  const entry = centralKey != null ? centralizedData[centralKey] : null;
  if (entry && typeof entry.description === 'string' && entry.description.trim()) {
    return entry.description;
  }
  return typeof photo.description === 'string' && photo.description ? photo.description : null;
}

/** The effective GPS position for a photo: the centralized
 *  metadata/photos.yaml position override if present, else the
 *  photo's real EXIF/OneDrive location, else null. Mirrors the
 *  override precedence in applyYamlFallbackFields in app.js. */
function effectivePosition(photo) {
  const centralKey = centralizedKeyFor(photo);
  const entry = centralKey != null ? centralizedData[centralKey] : null;
  const pos = entry?.position ?? {};
  const lat = pos.lat ?? pos.latitude;
  const lon = pos.long ?? pos.lon ?? pos.lng ?? pos.longitude;
  if (typeof lat === 'number' && typeof lon === 'number') {
    return { latitude: lat, longitude: lon };
  }

  const loc = photo.location;
  if (loc?.latitude != null && loc?.longitude != null) {
    return { latitude: loc.latitude, longitude: loc.longitude };
  }
  return null;
}

/** Recursively flattens marked.js inline tokens (text/strong/em/
 *  codespan/link/image/br) into a flat list of simple "runs" for
 *  rendering in jsPDF, which has no rich-text/HTML support of its
 *  own: { type: 'text', text, bold, italic, href } for words (href
 *  set if inside a Markdown link, so the whole link phrase renders
 *  clickable), or { type: 'image', src, alt } for an embedded image
 *  (e.g. a map screenshot), or { type: 'break' } for an explicit line
 *  break. */
function flattenInlineTokens(tokens, bold = false, italic = false) {
  const runs = [];
  for (const t of tokens || []) {
    switch (t.type) {
      case 'text':
      case 'escape':
        runs.push({ type: 'text', text: t.text, bold, italic });
        break;
      case 'strong':
        runs.push(...flattenInlineTokens(t.tokens, true, italic));
        break;
      case 'em':
        runs.push(...flattenInlineTokens(t.tokens, bold, true));
        break;
      case 'codespan':
        runs.push({ type: 'text', text: t.text, bold, italic });
        break;
      case 'link': {
        const inner = flattenInlineTokens(
          t.tokens && t.tokens.length ? t.tokens : [{ type: 'text', text: t.text }],
          bold,
          italic
        );
        inner.forEach((r) => {
          if (r.type === 'text') r.href = t.href;
        });
        runs.push(...inner);
        break;
      }
      case 'image':
        runs.push({ type: 'image', src: t.href, alt: t.text });
        break;
      case 'br':
        runs.push({ type: 'break' });
        break;
      default:
        if (t.tokens) runs.push(...flattenInlineTokens(t.tokens, bold, italic));
        else if (t.text) runs.push({ type: 'text', text: t.text, bold, italic });
    }
  }
  return runs;
}

/** Very small Markdown -> plain-text converter, used only as a
 *  fallback if marked.js somehow failed to load (renderMarkdownBlock
 *  normally handles gallery/day descriptions with full link/image
 *  support instead). Strips fenced code blocks, heading markers,
 *  bold/italic/code markers, and turns [text](url) links and list
 *  bullets into plain readable text (losing the link href and any
 *  images). */
function markdownToPlainText(md) {
  if (!md) return '';
  return md
    .replace(/```[\s\S]*?```/g, '')
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/!?\[([^\]]+)\]\([^)]+\)/g, '$1')
    .replace(/[*_`]{1,3}/g, '')
    .replace(/^\s*[-*+]\s+/gm, '• ')
    .replace(/^\s*\d+\.\s+/gm, '')
    .trim();
}

/** Matches a OneDrive/SharePoint "Anyone with the link" share URL
 *  (short 1drv.ms links or full onedrive.live.com/sharepoint.com
 *  links) -- these can't be fetch()'d directly for their image bytes:
 *  they resolve to an HTML viewer page, not raw image data, and
 *  typically don't send CORS headers anyway, so a direct fetch() then
 *  either downloads the wrong thing or fails outright. Used to embed
 *  a link like ![Kart](https://1drv.ms/i/...) that a user pastes into
 *  a gallery/day description's Markdown. */
function isOneDriveShareLink(url) {
  return /^https:\/\/(1drv\.ms|[\w-]+\.sharepoint\.com|onedrive\.live\.com)\//i.test(url);
}

/** Resolves a OneDrive/SharePoint share link to its actual direct
 *  download URL via the Graph `/shares` API (same encodeShareUrl
 *  encoding used for the gallery's own share link), using the
 *  signed-in user's access token. The resulting download URL is a
 *  temporary, pre-authenticated, CORS-enabled blob storage URL that
 *  can be fetched directly, unlike the original share link. */
async function resolveOneDriveShareLink(url) {
  if (!lastAccessToken) throw new Error('Not signed in, cannot resolve OneDrive link');
  const shareId = encodeShareUrl(url);
  // Deliberately no $select here: Graph silently omits the
  // @microsoft.graph.downloadUrl annotation from the response
  // whenever $select is used (a known Graph API quirk), so the full
  // driveItem must be requested to actually get it.
  const resp = await fetch(`${GRAPH_API}/shares/${shareId}/driveItem`, {
    headers: { Authorization: 'Bearer ' + lastAccessToken },
  });
  if (!resp.ok) throw new Error(`Could not resolve OneDrive link (HTTP ${resp.status})`);
  const data = await resp.json();
  const downloadUrl = data['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) throw new Error('OneDrive link has no direct download URL');
  return downloadUrl;
}

/** Fetches an image URL (OneDrive's pre-signed thumbnail CDN URLs
 *  work directly, no Authorization header needed) and converts it to
 *  a data URL for embedding via jsPDF's addImage. OneDrive/SharePoint
 *  "share" links (e.g. a 1drv.ms link pasted into a Markdown
 *  description) are resolved to their real direct-download URL first
 *  via the Graph API, since the share link itself isn't fetchable
 *  image data. */
async function fetchImageAsDataUrl(url) {
  const fetchUrl = isOneDriveShareLink(url) ? await resolveOneDriveShareLink(url) : url;
  const resp = await fetch(fetchUrl);
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const blob = await resp.blob();
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Could not read image data'));
    reader.readAsDataURL(blob);
  });
}

/** Loads a data URL into an <img> just to read its natural pixel
 *  dimensions, needed to scale it into the PDF page while preserving
 *  aspect ratio. */
function getImageDimensions(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => reject(new Error('Could not decode image'));
    img.src = dataUrl;
  });
}

/** Scales (w, h) down or up to fit within (maxW, maxH), preserving
 *  aspect ratio. */
function fitDimensions(w, h, maxW, maxH) {
  const ratio = Math.min(maxW / w, maxH / h);
  return { width: w * ratio, height: h * ratio };
}

/** jsPDF's addImage needs a format string ('JPEG'/'PNG'/etc.) matching
 *  the actual image bytes -- derive it from the data URL's mime type
 *  rather than assuming JPEG, in case OneDrive ever serves PNG
 *  thumbnails. */
function imageFormatFromDataUrl(dataUrl) {
  const match = dataUrl.match(/^data:image\/(\w+);/i);
  if (!match) return 'JPEG';
  const type = match[1].toUpperCase();
  return type === 'JPG' ? 'JPEG' : type;
}

/** Builds the full gallery PDF in memory and returns the jsPDF
 *  document (does not trigger the download itself -- see
 *  handleGeneratePdf). Structure: title page (heading + gallery
 *  description, rendered richly -- headings, bold/italic, clickable
 *  links, and embedded images all preserved) -> per date-group
 *  section heading + day description (same rich rendering) -> photos
 *  laid out 4-per-page in a 2x2 grid (image, filename, caption, taken
 *  date, GPS position + "View on map" link, each truncated to fit its
 *  compact cell) -> page numbers added in a final pass. Calls
 *  onProgress(current, total) before rendering each photo's cell so
 *  the caller can show progress. A photo whose thumbnail fails to
 *  fetch/decode still gets its caption cell, with an "(image
 *  unavailable)" note instead of aborting the whole export -- the
 *  same graceful fallback applies to any image embedded in a
 *  description. */
async function generateGalleryPdf(onProgress) {
  if (typeof window.jspdf === 'undefined') {
    throw new Error('PDF library failed to load (jsPDF).');
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 40;
  const contentWidth = pageWidth - margin * 2;
  let y = margin;

  function addWrappedText(text, fontSize) {
    if (!text) return;
    doc.setFontSize(fontSize);
    doc.setTextColor(0);
    const lines = doc.splitTextToSize(text, contentWidth);
    lines.forEach((line) => {
      if (y > pageHeight - margin) {
        doc.addPage();
        y = margin;
      }
      doc.text(line, margin, y);
      y += fontSize * 1.3;
    });
    y += 8;
  }

  /** Adds a new page and resets y to the top margin if the given
   *  height wouldn't fit on the current page before the bottom
   *  margin -- used throughout the rich Markdown renderer below so
   *  headings/paragraph lines/images never get cut off mid-block. */
  function ensureSpace(neededHeight) {
    if (y + neededHeight > pageHeight - margin) {
      doc.addPage();
      y = margin;
    }
  }

  const defaultFontName = doc.getFont().fontName;

  /** Renders a word-wrapped run of inline Markdown content (as
   *  produced by flattenInlineTokens) at the given base font size,
   *  preserving bold/italic styling, rendering Markdown links as
   *  clickable text (via doc.textWithLink), and embedding any inline
   *  images (e.g. a map screenshot) via doc.addImage -- flushing to a
   *  fresh line before/after each embedded image since it can't share
   *  a text line. An optional bulletPrefix (e.g. "• ") is prepended
   *  as a plain leading word, for list items. */
  async function renderInlineRuns(runs, fontSize, bulletPrefix = '') {
    const lineHeight = fontSize * 1.35;
    const spaceWidth = (() => {
      doc.setFont(defaultFontName, 'normal');
      doc.setFontSize(fontSize);
      return doc.getTextWidth(' ');
    })();

    const words = [];
    if (bulletPrefix) words.push({ type: 'text', text: bulletPrefix, bold: false, italic: false });
    for (const run of runs) {
      if (run.type === 'image' || run.type === 'break') {
        words.push(run);
        continue;
      }
      for (const part of run.text.split(/\s+/)) {
        if (part) words.push({ type: 'text', text: part, bold: run.bold, italic: run.italic, href: run.href });
      }
    }

    let x = margin;
    ensureSpace(lineHeight);

    for (const word of words) {
      if (word.type === 'image') {
        x = margin;
        y += 4;
        try {
          const dataUrl = await fetchImageAsDataUrl(word.src);
          const dims = await getImageDimensions(dataUrl);
          const maxImgHeight = 220;
          const { width, height } = fitDimensions(dims.width, dims.height, contentWidth, maxImgHeight);
          ensureSpace(height + 10);
          doc.addImage(
            dataUrl,
            imageFormatFromDataUrl(dataUrl),
            margin + (contentWidth - width) / 2,
            y,
            width,
            height
          );
          y += height + 10;
        } catch (err) {
          // Fall through -- show a small note instead of the image
          // rather than dropping/aborting the whole description. Log
          // the real cause to the console since the PDF note alone
          // ("[image unavailable: ...]") isn't enough to debug why
          // (e.g. not signed in, share link needs resolving, CORS,
          // 404, etc.).
          console.warn(`PDF export: could not embed description image "${word.alt || word.src}" (${word.src}):`, err);
          ensureSpace(lineHeight);
          doc.setFont(defaultFontName, 'normal');
          doc.setFontSize(fontSize - 1);
          doc.setTextColor(150);
          doc.text(`[image unavailable: ${word.alt || word.src}]`, margin, y);
          doc.setTextColor(0);
          y += lineHeight;
        }
        continue;
      }
      if (word.type === 'break') {
        x = margin;
        y += lineHeight;
        ensureSpace(lineHeight);
        continue;
      }

      const style = word.bold && word.italic ? 'bolditalic' : word.bold ? 'bold' : word.italic ? 'italic' : 'normal';
      doc.setFont(defaultFontName, style);
      doc.setFontSize(fontSize);
      const wordWidth = doc.getTextWidth(word.text);

      if (x !== margin && x + wordWidth > margin + contentWidth) {
        x = margin;
        y += lineHeight;
        ensureSpace(lineHeight);
      }

      if (word.href) {
        doc.setTextColor(40, 90, 200);
        doc.textWithLink(word.text, x, y, { url: word.href });
        doc.setTextColor(0);
      } else {
        doc.text(word.text, x, y);
      }
      x += wordWidth + spaceWidth;
    }
    doc.setFont(defaultFontName, 'normal');
    y += lineHeight;
  }

  /** Renders one marked.js block-level token (heading, paragraph,
   *  list, image, hr, etc.) at the given base font size. */
  async function renderMarkdownBlockToken(block, fontSize) {
    switch (block.type) {
      case 'heading': {
        const size = Math.max(fontSize, 18 - (block.depth - 1) * 2);
        ensureSpace(size * 1.4);
        doc.setFont(defaultFontName, 'bold');
        doc.setFontSize(size);
        doc.setTextColor(0);
        doc.text(block.text, margin, y);
        doc.setFont(defaultFontName, 'normal');
        y += size * 1.4 + 4;
        break;
      }
      case 'paragraph':
        await renderInlineRuns(flattenInlineTokens(block.tokens), fontSize);
        y += 6;
        break;
      case 'list':
        for (const item of block.items) {
          await renderInlineRuns(flattenInlineTokens(item.tokens), fontSize, '• ');
        }
        y += 4;
        break;
      case 'hr':
        ensureSpace(14);
        doc.setDrawColor(200);
        doc.line(margin, y, margin + contentWidth, y);
        doc.setDrawColor(0);
        y += 14;
        break;
      case 'space':
        break;
      case 'code':
        addWrappedText(block.text, fontSize - 1);
        break;
      default:
        if (block.tokens) {
          await renderInlineRuns(flattenInlineTokens(block.tokens), fontSize);
        } else if (block.text) {
          addWrappedText(block.text, fontSize);
        }
    }
  }

  /** Renders a full Markdown string (the gallery description or a
   *  day description) with headings, bold/italic, clickable links,
   *  list bullets, and embedded images (e.g. a map screenshot) all
   *  preserved -- unlike plain-text stripping, which loses link hrefs
   *  and drops images entirely. Falls back to plain-text rendering if
   *  marked.js somehow isn't available. */
  async function renderMarkdown(markdownSrc, fontSize) {
    if (!markdownSrc || !markdownSrc.trim()) return;
    if (typeof marked === 'undefined' || typeof marked.lexer !== 'function') {
      addWrappedText(markdownToPlainText(markdownSrc), fontSize);
      return;
    }
    const blocks = marked.lexer(markdownSrc);
    for (const block of blocks) {
      await renderMarkdownBlockToken(block, fontSize);
    }
  }

  // Title page
  doc.setFontSize(22);
  doc.text('Photo Gallery', margin, y + 10);
  y += 34;
  doc.setFontSize(10);
  doc.setTextColor(120);
  doc.text('Generated ' + new Date().toLocaleDateString(), margin, y);
  doc.setTextColor(0);
  y += 26;
  await renderMarkdown(galleryDescriptionText, 11);

  // Photos are laid out 4-per-page in a 2x2 grid (one row of 2, then
  // another row of 2), each cell holding a scaled-to-fit image plus a
  // compact single-line-per-field caption/date/GPS block below it.
  const gridCols = 2;
  const gridRows = 2;
  const cellGapX = 18;
  const cellGapY = 24;
  const cellWidth = (contentWidth - cellGapX * (gridCols - 1)) / gridCols;
  const gridTop = margin;
  const gridBottom = pageHeight - margin;
  const cellHeight = (gridBottom - gridTop - cellGapY * (gridRows - 1)) / gridRows;
  const textBlockHeight = 58; // reserved for filename/caption/date/GPS lines
  const imageMaxHeight = cellHeight - textBlockHeight - 6;

  /** Truncates text to a single line that fits maxWidth at the given
   *  font size, appending an ellipsis if it had to be cut short --
   *  used for the compact per-cell caption fields since each cell has
   *  very little vertical space in a 4-per-page grid. */
  function truncateLine(text, fontSize, maxWidth) {
    doc.setFontSize(fontSize);
    const lines = doc.splitTextToSize(text, maxWidth);
    if (lines.length <= 1) return lines[0] || '';
    let line = lines[0];
    while (line.length > 1 && doc.getTextWidth(line + '…') > maxWidth) {
      line = line.slice(0, -1);
    }
    return line.replace(/\s+$/, '') + '…';
  }

  const groups = groupPhotosByDate(photos);
  const totalPhotos = photos.length;
  let photoIndex = 0;

  for (const group of groups) {
    doc.addPage();
    y = margin;
    doc.setFontSize(16);
    doc.text(group.label, margin, y);
    y += 24;

    const dayText = group.key !== 'unknown' ? dayDescriptionTexts[group.key] : null;
    await renderMarkdown(dayText, 11);

    let cellIndex = 0;
    for (const photo of group.items) {
      photoIndex++;
      if (onProgress) onProgress(photoIndex, totalPhotos);

      if (cellIndex === 0) doc.addPage();
      const row = Math.floor(cellIndex / gridCols);
      const col = cellIndex % gridCols;
      const cellX = margin + col * (cellWidth + cellGapX);
      const cellY = gridTop + row * (cellHeight + cellGapY);
      cellIndex = (cellIndex + 1) % (gridCols * gridRows);

      const thumbUrl =
        photo.thumbnails?.[0]?.large?.url || photo.thumbnails?.[0]?.medium?.url || '';
      let imageDrawn = false;
      if (thumbUrl) {
        try {
          const dataUrl = await fetchImageAsDataUrl(thumbUrl);
          const dims = await getImageDimensions(dataUrl);
          const { width, height } = fitDimensions(
            dims.width,
            dims.height,
            cellWidth,
            imageMaxHeight
          );
          doc.addImage(
            dataUrl,
            imageFormatFromDataUrl(dataUrl),
            cellX + (cellWidth - width) / 2,
            cellY + (imageMaxHeight - height) / 2,
            width,
            height
          );
          imageDrawn = true;
        } catch {
          // Fall through -- render the caption block without the
          // image rather than aborting the whole export.
        }
      }
      if (!imageDrawn) {
        doc.setFontSize(9);
        doc.setTextColor(150);
        doc.text('(image unavailable)', cellX, cellY + imageMaxHeight / 2);
        doc.setTextColor(0);
      }

      let textY = cellY + imageMaxHeight + 12;
      doc.setFontSize(9);
      doc.setTextColor(0);
      doc.text(truncateLine(photo.name, 9, cellWidth), cellX, textY);
      textY += 12;

      const desc = effectiveDescription(photo);
      if (desc) {
        doc.setFontSize(8);
        doc.setTextColor(60);
        doc.text(truncateLine(desc, 8, cellWidth), cellX, textY);
        doc.setTextColor(0);
        textY += 11;
      }

      const taken = effectiveTakenDate(photo);
      if (taken && !isNaN(taken)) {
        doc.setFontSize(8);
        doc.text(truncateLine('Taken: ' + taken.toLocaleString(), 8, cellWidth), cellX, textY);
        textY += 11;
      }

      const pos = effectivePosition(photo);
      if (pos) {
        doc.setFontSize(8);
        const posText = `${pos.latitude.toFixed(5)}, ${pos.longitude.toFixed(5)} — `;
        doc.text(posText, cellX, textY);
        const mapUrl = `https://www.google.com/maps?q=${pos.latitude},${pos.longitude}`;
        doc.setTextColor(40, 90, 200);
        doc.textWithLink('View on map', cellX + doc.getTextWidth(posText), textY, {
          url: mapUrl,
        });
        doc.setTextColor(0);
      }
    }
  }

  const totalPages = doc.internal.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setFontSize(9);
    doc.setTextColor(150);
    doc.text(`Page ${i} of ${totalPages}`, pageWidth - margin - 70, pageHeight - 20);
    doc.setTextColor(0);
  }

  return doc;
}

function pdfStatus(message, kind) {
  const el = document.getElementById('admin-pdf-status');
  el.textContent = message;
  el.className = 'admin-status' + (kind ? ` ${kind}` : '');
  el.hidden = !message;
}

async function handleGeneratePdf() {
  const btn = document.getElementById('admin-pdf-generate-btn');
  btn.disabled = true;
  pdfStatus('Preparing…', '');
  try {
    const doc = await generateGalleryPdf((current, total) => {
      pdfStatus(`Rendering photo ${current} of ${total}…`, '');
    });
    doc.save('gallery.pdf');
    pdfStatus('PDF downloaded ✓', 'success');
  } catch (err) {
    pdfStatus('Failed to generate PDF: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
  }
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

    galleryDescriptionText = await fetchGalleryDescription(result.items, token);
    lastSavedDescriptionText = galleryDescriptionText;
    document.getElementById('admin-description-text').value = galleryDescriptionText;
    renderDescriptionPreview();

    const dayMap = await fetchDayDescriptions(result.items, token);
    dayDescriptionTexts = Object.fromEntries(dayMap);
    dayDescriptionDates = computeDayDescriptionDates();
    selectedDayKey = null;
    document.getElementById('admin-daydesc-empty').hidden = false;
    document.getElementById('admin-daydesc-content').hidden = true;
    renderDayDescriptionList();

    document.getElementById('admin-tabs').hidden = false;
    switchTab(activeTab);
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
document.getElementById('admin-description-text').addEventListener('input', () => {
  renderDescriptionPreview();
  descriptionStatus('');
});
document.getElementById('admin-description-save-btn').addEventListener('click', handleSaveDescription);
document.getElementById('admin-daydesc-text').addEventListener('input', () => {
  renderDayDescPreview();
  dayDescStatus('');
});
document.getElementById('admin-daydesc-save-btn').addEventListener('click', handleSaveDayDescription);
document.getElementById('admin-pdf-generate-btn').addEventListener('click', handleGeneratePdf);
document.querySelectorAll('.admin-tab-btn').forEach((btn) => {
  btn.addEventListener('click', () => switchTab(btn.dataset.tab));
});
window.addEventListener('beforeunload', (e) => {
  const descChanged =
    document.getElementById('admin-description-text')?.value !== lastSavedDescriptionText;
  const dayChanged =
    selectedDayKey != null &&
    document.getElementById('admin-daydesc-text')?.value !== lastSavedDayText;
  if (descChanged || dayChanged) {
    e.preventDefault();
    e.returnValue = '';
  }
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
