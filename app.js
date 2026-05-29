// ── Config ───────────────────────────────────────────────────────────────────
const CONFIG = {
  clientId: '64f2d2d0-6c34-4af1-95ec-114f230be16b',
  authority: 'https://login.microsoftonline.com/common',
  redirectUri: 'https://hemantthakkar.github.io/slide-addin/',
  scopes: ['Files.Read', 'Files.Read.All', 'Sites.Read.All', 'User.Read']
};

// ── State ────────────────────────────────────────────────────────────────────
let msalInstance = null;
let accessToken  = null;
let allFiles     = [];
let currentFile  = null;
let currentSlides = [];
let siteBarVisible = false;
const LS = localStorage;

// ── Wait for everything to load ──────────────────────────────────────────────
function waitForMsal(callback) {
  if (typeof msal !== 'undefined') {
    callback();
  } else {
    setTimeout(() => waitForMsal(callback), 100);
  }
}

function initMsal() {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: CONFIG.authority,
      redirectUri: CONFIG.redirectUri
    },
    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
  });
}

// ── Init ─────────────────────────────────────────────────────────────────────
function appInit() {
  waitForMsal(() => {
    initMsal();
    tryAutoLogin();
  });
}

// Try Office.onReady first, fall back to window.onload
if (typeof Office !== 'undefined') {
  Office.onReady(() => appInit());
} else {
  window.addEventListener('load', () => appInit());
}

// ── Auto Login ───────────────────────────────────────────────────────────────
async function tryAutoLogin() {
  try {
    await msalInstance.handleRedirectPromise();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      const result = await msalInstance.acquireTokenSilent({
        scopes: CONFIG.scopes,
        account: accounts[0]
      });
      accessToken = result.accessToken;
      setUser(accounts[0].name || accounts[0].username);
      showApp();
    }
  } catch(e) { /* not logged in yet, that's fine */ }
}

// ── Login / Logout ───────────────────────────────────────────────────────────
async function login() {
  if (!msalInstance) {
    showError('Auth library not ready. Please wait a moment and try again.');
    return;
  }
  try {
    const result = await msalInstance.loginPopup({ scopes: CONFIG.scopes });
    accessToken = result.accessToken;
    setUser(result.account.name || result.account.username);
    hideError();
    showApp();
  } catch(e) {
    showError('Sign-in failed: ' + (e.message || 'Please try again'));
  }
}

function logout() {
  accessToken = null; allFiles = []; currentFile = null;
  try { msalInstance.logoutPopup(); } catch(e) {}
  showScreen('login-screen');
  document.getElementById('header-user').style.display = 'none';
  document.getElementById('btn-logout').style.display = 'none';
}

function setUser(name) {
  const short = (name || '').split(' ')[0] || name;
  document.getElementById('header-user').textContent = short;
  document.getElementById('header-user').style.display = 'block';
  document.getElementById('btn-logout').style.display = 'block';
}

// ── Screen Management ────────────────────────────────────────────────────────
function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}

function showApp() {
  showScreen('app-screen');
  const saved = LS.getItem('sl_site_url');
  if (saved) document.getElementById('site-url').value = saved;
  loadFiles();
}

function showFiles() {
  document.getElementById('files-view').style.display = 'block';
  document.getElementById('slides-view').classList.remove('active');
  document.getElementById('main-search').value = '';
  filterFiles();
}

function showSlides() {
  document.getElementById('files-view').style.display = 'none';
  document.getElementById('slides-view').classList.add('active');
}

// ── API ──────────────────────────────────────────────────────────────────────
async function apiGet(url) {
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + accessToken } });
  if (r.status === 401) { accessToken = null; logout(); throw new Error('Session expired'); }
  if (!r.ok) throw new Error(`API ${r.status}`);
  return r.json();
}

// ── Load Files ───────────────────────────────────────────────────────────────
async function loadFiles() {
  setStatus('<span class="spinner"></span> Loading files…', true);
  document.getElementById('files-view').innerHTML = '';
  try {
    let items = [];
    const siteUrl = LS.getItem('sl_site_url');
    if (siteUrl) {
      try {
        const u = new URL(siteUrl);
        const siteData = await apiGet(`https://graph.microsoft.com/v1.0/sites/${u.hostname}:${u.pathname}`);
        const drives = await apiGet(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/drives`);
        for (const drive of drives.value.slice(0, 5)) {
          try {
            const res = await apiGet(`https://graph.microsoft.com/v1.0/drives/${drive.id}/root/search(q='.pptx')?$top=50`);
            items = items.concat(res.value || []);
          } catch(e) {}
        }
      } catch(e) {
        setStatus('⚠️ Could not access SharePoint site. Check the URL.'); return;
      }
    } else {
      const res = await apiGet(`https://graph.microsoft.com/v1.0/me/drive/root/search(q='.pptx')?$top=100`);
      items = res.value || [];
    }
    allFiles = items.filter(i => i.name && i.name.toLowerCase().endsWith('.pptx'));
    renderFiles(allFiles);
    setStatus(`${allFiles.length} file${allFiles.length !== 1 ? 's' : ''} found`);
  } catch(e) {
    setStatus('⚠️ Error: ' + e.message);
  }
}

function setStatus(html) {
  document.getElementById('status').innerHTML = html;
}

function renderFiles(files) {
  const view = document.getElementById('files-view');
  if (!files.length) {
    view.innerHTML = `<div class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 13h6m-3-3v6m-9 1V7a2 2 0 012-2h6l2 2h6a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg><p>No PowerPoint files found.<br>Try setting a SharePoint site URL in Settings.</p></div>`;
    return;
  }
  view.innerHTML = files.map(f => {
    const modified = f.lastModifiedDateTime ? new Date(f.lastModifiedDateTime).toLocaleDateString() : '';
    const size = f.size ? (f.size > 1e6 ? (f.size/1e6).toFixed(1)+' MB' : Math.round(f.size/1e3)+' KB') : '';
    return `<div class="file-item" onclick='openFile(${JSON.stringify(f).replace(/'/g,"&#39;")})'>
      <div class="file-badge">PPTX</div>
      <div class="file-details">
        <div class="file-name" title="${f.name}">${f.name.replace(/\.pptx$/i,'')}</div>
        <div class="file-meta">${[modified, size].filter(Boolean).join(' · ')}</div>
      </div>
      <div class="file-arrow"><svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clip-rule="evenodd"/></svg></div>
    </div>`;
  }).join('');
}

function filterFiles() {
  const q = document.getElementById('main-search').value.toLowerCase();
  renderFiles(allFiles.filter(f => f.name.toLowerCase().includes(q)));
}

// ── Open File ────────────────────────────────────────────────────────────────
async function openFile(file) {
  currentFile = file;
  currentSlides = [];
  document.getElementById('slides-title').textContent = file.name.replace(/\.pptx$/i,'');
  document.getElementById('slide-search').value = '';
  showSlides();
  document.getElementById('slides-grid').innerHTML = `<div style="grid-column:1/-1;text-align:center;padding:20px;color:var(--text3)"><span class="spinner" style="display:inline-block"></span></div>`;

  const count = Math.min(Math.max(1, Math.floor((file.size || 50000) / 35000)), 40);
  const colors = ['#E8F0FE','#FEF3E2','#E6F4EA','#FCE8E6','#E8EAF6','#F3E5F5','#E0F7FA','#FFF3E0'];
  currentSlides = Array.from({length: count}, (_, i) => ({ index: i+1, color: colors[i % colors.length] }));

  try {
    const driveId = file.parentReference?.driveId;
    const ep = driveId
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${file.id}/thumbnails?$select=large`
      : `https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/thumbnails?$select=large`;
    const td = await apiGet(ep);
    if (td.value?.[0]?.large?.url) {
      currentSlides[0].thumb = td.value[0].large.url;
    }
  } catch(e) {}

  renderSlides(currentSlides);
}

function renderSlides(slides) {
  const grid = document.getElementById('slides-grid');
  if (!slides.length) {
    grid.innerHTML = `<div style="grid-column:1/-1" class="empty"><p>No slides found</p></div>`;
    return;
  }
  grid.innerHTML = slides.map(s => {
    const thumbContent = s.thumb
      ? `<img src="${s.thumb}" style="width:100%;height:100%;object-fit:cover;position:absolute;top:0;left:0;" />`
      : `<span style="font-size:10px;color:#3d5a99;font-weight:600;padding:8px;text-align:center;line-height:1.3;z-index:1;">Slide ${s.index}</span>`;
    return `<div class="slide-card">
      <div class="slide-thumb" style="background:${s.color};position:relative;">
        <span class="slide-num-badge">${s.index}</span>
        ${thumbContent}
      </div>
      <div class="slide-footer">
        <span class="slide-label">Slide ${s.index}</span>
        <button class="btn-insert" onclick="insertSlideNote(${s.index})">Use</button>
      </div>
    </div>`;
  }).join('');
}

function filterSlides() {
  const q = document.getElementById('slide-search').value.toLowerCase();
  renderSlides(currentSlides.filter(s => String(s.index).includes(q)));
}

// ── Actions ──────────────────────────────────────────────────────────────────
function insertSlideNote(idx) {
  showToast(`Open file → copy slide ${idx} → paste here`);
  if (currentFile?.webUrl) {
    try { Office.context.ui.openBrowserWindow(currentFile.webUrl); } catch(e) {
      window.open(currentFile.webUrl, '_blank');
    }
  }
}

async function downloadFile() {
  if (!currentFile) return;
  try {
    const driveId = currentFile.parentReference?.driveId;
    const ep = driveId
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${currentFile.id}`
      : `https://graph.microsoft.com/v1.0/me/drive/items/${currentFile.id}`;
    const data = await apiGet(ep);
    if (data['@microsoft.graph.downloadUrl']) {
      try { Office.context.ui.openBrowserWindow(data['@microsoft.graph.downloadUrl']); } catch(e) {
        window.open(data['@microsoft.graph.downloadUrl'], '_blank');
      }
    } else showToast('Download link unavailable');
  } catch(e) { showToast('Error: ' + e.message); }
}

function openInSharePoint() {
  if (currentFile?.webUrl) {
    try { Office.context.ui.openBrowserWindow(currentFile.webUrl); } catch(e) {
      window.open(currentFile.webUrl, '_blank');
    }
  } else showToast('URL not available');
}

// ── Site Bar ─────────────────────────────────────────────────────────────────
function toggleSiteBar() {
  siteBarVisible = !siteBarVisible;
  document.getElementById('site-bar').classList.toggle('visible', siteBarVisible);
}

function applySite() {
  const url = document.getElementById('site-url').value.trim();
  LS.setItem('sl_site_url', url);
  siteBarVisible = false;
  document.getElementById('site-bar').classList.remove('visible');
  showFiles();
  loadFiles();
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3000);
}

function showError(msg) {
  const el = document.getElementById('login-error');
  el.textContent = msg;
  el.style.display = 'block';
}

function hideError() {
  document.getElementById('login-error').style.display = 'none';
}
