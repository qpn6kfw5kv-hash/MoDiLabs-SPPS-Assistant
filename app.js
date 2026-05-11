// ============================================================
// MoDiLabs SPPS Lab Assistant — Main Application Logic (v3.1)
// Features: Folder system, Tabs, global search, multiple activators
// ============================================================

// ---- State Management ----
const APP_KEY = 'spps-lab-assistant-data';
const FILES_KEY = 'spps-lab-assistant-files';
const FOLDERS_KEY = 'spps-lab-assistant-folders';

function loadSyntheses() { try { return JSON.parse(localStorage.getItem(APP_KEY)) || []; } catch { return []; } }
function saveSyntheses(data, skipWorkspace = false) { 
  localStorage.setItem(APP_KEY, JSON.stringify(data)); 
  if (workspaceHandle && !skipWorkspace) saveToWorkspace();
}

function loadFiles() { try { return JSON.parse(localStorage.getItem(FILES_KEY)) || {}; } catch { return {}; } }
function saveFiles(data, skipWorkspace = false) { 
  try { 
    localStorage.setItem(FILES_KEY, JSON.stringify(data)); 
    if (workspaceHandle && !skipWorkspace) saveToWorkspace();
  } catch { alert('⚠️ Spazio di archiviazione pieno.'); } 
}

function loadFolders() { try { return JSON.parse(localStorage.getItem(FOLDERS_KEY)) || []; } catch { return []; } }
function saveFolders(data, skipWorkspace = false) { 
  localStorage.setItem(FOLDERS_KEY, JSON.stringify(data)); 
  if (workspaceHandle && !skipWorkspace) saveToWorkspace();
}

function generateId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

// ---- App State ----
// ---- App State ----
const IS_SECURE_CONTEXT = window.isSecureContext || window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
let syntheses = loadSyntheses();
let attachments = loadFiles();
let folders = loadFolders();
let inventoryData = [];
let inventoryFileName = null;
let workspaceHandle = null;
let workspaceName = null;
let currentView = 'dashboard'; 
let currentTab = 'folders'; 
let currentFolderId = null;
let currentSynthesisId = null;
let searchQuery = '';

// ---- Inventory / Excel Logic ----
async function initInventory() {
  if (!IS_SECURE_CONTEXT) {
    console.warn('⚠️ L\'accesso al file system richiede un contesto sicuro (HTTPS o localhost).');
    return;
  }
  try {
    const invHandle = await idbKeyval.get('spps-inventory-handle');
    if (invHandle && (await verifyPermission(invHandle, false))) {
       await parseExcelHandle(invHandle);
    }
    const wsHandle = await idbKeyval.get('spps-workspace-handle');
    if (wsHandle && (await verifyPermission(wsHandle, true))) {
       workspaceHandle = wsHandle;
       workspaceName = wsHandle.name;
       await syncFromWorkspace();
    }
  } catch (err) { console.error('Init error', err); }
}

async function verifyPermission(fileHandle, readWrite) {
  const opts = { mode: readWrite ? 'readwrite' : 'read' };
  if ((await fileHandle.queryPermission(opts)) === 'granted') return true;
  if ((await fileHandle.requestPermission(opts)) === 'granted') return true;
  return false;
}

async function connectInventory() {
  try {
    const [fileHandle] = await window.showOpenFilePicker({
      types: [{ description: 'Excel Files', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls']} }]
    });
    await idbKeyval.set('spps-inventory-handle', fileHandle);
    await parseExcelHandle(fileHandle);
    
    // If workspace is active, register this filename for others
    if (workspaceHandle) {
      // We don't save the handle but the name, expecting the file to be in the workspace folder
      await saveToWorkspace(); 
    }
    render();
  } catch(e) { console.error(e); }
}

async function parseExcelHandle(handle) {
  try {
    const file = await handle.getFile();
    await parseExcelFile(file);
  } catch (err) { console.error("Error reading excel handle:", err); }
}

async function parseExcelFile(file) {
  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    inventoryData = rawData.filter(row => row && row.length > 0);
    inventoryFileName = file.name;
    render();
  } catch (err) {
    console.error("Error parsing excel file:", err);
    inventoryData = [];
    inventoryFileName = null;
  }
}

async function connectWorkspace() {
  try {
    const handle = await window.showDirectoryPicker();
    await idbKeyval.set('spps-workspace-handle', handle);
    workspaceHandle = handle;
    workspaceName = handle.name;
    // Try to load existing data from folder
    await syncFromWorkspace();
    render();
  } catch(e) { console.error(e); }
}

async function syncFromWorkspace() {
  if (!workspaceHandle) return;
  try {
    const fileHandle = await workspaceHandle.getFileHandle('spps_data.json', { create: true });
    const file = await fileHandle.getFile();
    const text = await file.text();
    if (text) {
      const data = JSON.parse(text);
      if (data.syntheses) { syntheses = data.syntheses; saveSyntheses(syntheses, true); }
      if (data.folders) { folders = data.folders; saveFolders(folders, true); }
      if (data.attachments) { attachments = data.attachments; saveFiles(attachments, true); }
      
      // Auto-load inventory from workspace if specified
      if (data.inventoryFileName && !inventoryData.length) {
        try {
          const invHandle = await workspaceHandle.getFileHandle(data.inventoryFileName);
          await parseExcelHandle(invHandle);
        } catch (e) { console.log("Workspace inventory not found or inaccessible:", data.inventoryFileName); }
      }
    }
  } catch (err) { console.error("Sync read error", err); }
}

async function saveToWorkspace() {
  if (!workspaceHandle) return;
  try {
    const data = { 
      syntheses, 
      folders, 
      attachments,
      inventoryFileName: inventoryFileName // Save current inventory name to shared project
    };
    const fileHandle = await workspaceHandle.getFileHandle('spps_data.json', { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(JSON.stringify(data, null, 2));
    await writable.close();
    showSaved('workspace-sync-status');
  } catch (err) { console.error("Sync write error", err); }
}

function exportData() {
  const data = { syntheses, folders, attachments, exportedAt: new Date().toISOString() };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `spps_backup_${new Date().toISOString().split('T')[0]}.json`;
  a.click();
}

function importData() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';
  input.onchange = async e => {
    const file = e.target.files[0];
    if (!file) return;
    const text = await file.text();
    try {
      const data = JSON.parse(text);
      if (confirm("Importare i dati? Questo sovrascriverà le sintesi locali.")) {
        if (data.syntheses) { syntheses = data.syntheses; saveSyntheses(syntheses); }
        if (data.folders) { folders = data.folders; saveFolders(folders); }
        if (data.attachments) { attachments = data.attachments; saveFiles(attachments); }
        render();
      }
    } catch(e) { alert("Errore nel file JSON"); }
  };
  input.click();
}


// ---- Router ----
function navigate(view, id = null) {
  currentView = view;
  if (view === 'dashboard') { currentFolderId = null; currentSynthesisId = null; }
  else if (view === 'folder') { currentFolderId = id; currentSynthesisId = null; }
  else if (view === 'detail') { currentSynthesisId = id; }
  searchQuery = '';
  render();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ---- Render Engine ----
function render() {
  const app = document.getElementById('app');
  if (currentView === 'dashboard') {
    app.innerHTML = renderHeader() + renderDashboard();
    bindDashboardEvents();
  } else if (currentView === 'folder') {
    const folder = folders.find(f => f.id === currentFolderId);
    if (!folder) { navigate('dashboard'); return; }
    app.innerHTML = renderHeader() + renderFolderView(folder);
    bindFolderViewEvents(folder);
  } else if (currentView === 'detail') {
    const synthesis = syntheses.find(s => s.id === currentSynthesisId);
    if (!synthesis) { navigate('dashboard'); return; }
    app.innerHTML = renderHeader() + renderDetail(synthesis);
    bindDetailEvents(synthesis);
  }
}

// ============================================================
// HEADER — Centered with MoDiLabs branding
// ============================================================
function renderHeader() {
  return `
    <header class="app-header-centered">
      <div class="app-brand">
        <img src="logo_transparent.png" alt="MoDiLabs Logo" class="app-logo-img">
        <div class="app-logo-title">SPPS Lab Assistant</div>
      </div>
    </header>
  `;
}

// ============================================================
// DASHBOARD
// ============================================================
function renderDashboard() {
  const stats = getStats();
  const isFoldersTab = currentTab === 'folders';

  return `
    <div class="stats-bar animate-in">
      <div class="stat-card">
        <div class="stat-icon total">🧪</div>
        <div><div class="stat-value">${stats.total}</div><div class="stat-label">Sintesi Totali</div></div>
      </div>
      <div class="stat-card">
        <div class="stat-icon" style="background:rgba(139,92,246,0.15); color:var(--purple)">📁</div>
        <div><div class="stat-value">${folders.length}</div><div class="stat-label">Cartelle</div></div>
      </div>
      <div class="stat-card">
        <div class="stat-icon progress">⏳</div>
        <div><div class="stat-value">${stats.inProgress}</div><div class="stat-label">In Corso</div></div>
      </div>
      <div class="stat-card">
        <div class="stat-icon analyzed">📊</div>
        <div><div class="stat-value">${stats.analyzed}</div><div class="stat-label">Analizzate</div></div>
      </div>
    </div>

    <!-- INVENTORY CARD -->
    <div class="dashboard-toolbar animate-in" style="margin-top: 24px; margin-bottom: 24px; padding: 16px 20px; background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border); box-shadow: var(--shadow-sm); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
      <div style="flex: 1; min-width: 250px;">
        <h3 style="margin:0; font-size: 1.1rem; color: var(--text-primary); display: flex; align-items: center; gap: 8px;">📦 Magazzino / Inventario Excel</h3>
        <p style="margin: 4px 0 0; font-size: 0.85rem; color: var(--text-muted);">${inventoryFileName ? 'Connesso a: <strong>' + escapeHtml(inventoryFileName) + '</strong> (' + inventoryData.length + ' righe mappate)' : 'Verrà sincronizzato in locale tramite OneDrive. Clicca per autorizzare la lettura.'}</p>
      </div>
      <button class="btn btn-primary" onclick="connectInventory()" style="background: ${inventoryFileName ? 'var(--info)' : 'var(--success)'}; color: white; border: none;">
        ${inventoryFileName ? '🔄 Sincronizza Dati' : '🔗 Collega Inventario Excel'}
      </button>
    </div>

    <!-- WORKSPACE CARD -->
    <div class="dashboard-toolbar animate-in" style="margin-top: 24px; margin-bottom: 24px; padding: 16px 20px; background: rgba(139,92,246,0.06); border-radius: var(--radius-md); border: 1px solid var(--purple-glow); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
      <div style="flex: 1; min-width: 250px;">
        <h3 style="margin:0; font-size: 1.1rem; color: var(--text-primary); display: flex; align-items: center; gap: 8px;">🏢 Area di Lavoro Lab (Shared)</h3>
        <p style="margin: 4px 0 0; font-size: 0.85rem; color: var(--text-muted);">${workspaceName ? 'Sincronizzato con cartella: <strong>' + escapeHtml(workspaceName) + '</strong>' : 'Sincronizza le tue sintesi su una cartella condivisa (OneDrive/Teams) per collaborare con i colleghi.'}</p>
        <div id="workspace-sync-status" class="notes-saved" style="margin-top:4px">✓ Dati sincronizzati sul server lab</div>
      </div>
      <div style="display:flex; gap:8px;">
        <button class="btn btn-secondary btn-sm" onclick="exportData()" title="Scarica backup locale">💾 Backup</button>
        <button class="btn btn-secondary btn-sm" onclick="importData()" title="Carica dati da file">📂 Importa</button>
        <button class="btn btn-primary" onclick="connectWorkspace()" style="background: var(--purple); color: white; border: none;">
          ${workspaceName ? '🔄 Riconnetti Area Lab' : '🔗 Attiva Area Condivisa'}
        </button>
      </div>
    </div>

    <!-- TABS -->
    <div class="dashboard-tabs animate-in">
      <button class="tab-btn ${isFoldersTab ? 'active' : ''}" data-tab="folders">📁 Cartelle Progetto</button>
      <button class="tab-btn ${!isFoldersTab ? 'active' : ''}" data-tab="files">📄 Tutte le Sequenze</button>
    </div>

    ${isFoldersTab ? renderFoldersTab() : renderAllFilesTab()}

    ${renderNewFolderModal()}
    ${renderNewSynthesisModal(null)}
  `;
}

// --- FOLDERS TAB ---
function renderFoldersTab() {
  return `
    <div class="dashboard-toolbar animate-in" style="margin-top: 20px;">
      <h2 class="section-heading">Gestione Cartelle</h2>
      <button class="btn btn-primary" id="btn-new-folder"><span>＋</span> Nuova Cartella</button>
    </div>
    
    <div class="search-bar animate-in">
      <span class="search-icon">🔍</span>
      <input type="text" id="search-folder-input" placeholder="Cerca una cartella..." value="${searchQuery}">
    </div>

    <div id="folders-container">
      ${renderFoldersGridList()}
    </div>
  `;
}

function renderFoldersGridList() {
  const filtered = searchQuery ? folders.filter(f => f.name.toLowerCase().includes(searchQuery.toLowerCase())) : folders;
  if (folders.length === 0) return renderFoldersEmpty();
  if (filtered.length === 0) return `<div class="empty-state animate-in"><div class="empty-state-icon">🔍</div><h3>Nessuna cartella trovata</h3></div>`;

  const FOLDER_COLORS = ['#06b6d4', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444', '#ec4899', '#3b82f6', '#14b8a6'];
  return `
    <div class="folders-grid animate-in">
      ${filtered.map(f => {
        const count = syntheses.filter(s => s.folderId === f.id).length;
        const color = f.color || FOLDER_COLORS[0];
        return `
          <div class="folder-card" data-id="${f.id}" style="--folder-color: ${color}">
            <div class="folder-card-icon">📁</div>
            <div class="folder-card-body">
              <h3 class="folder-card-name">${escapeHtml(f.name)}</h3>
              <span class="folder-card-count">${count} ${count === 1 ? 'sequenza' : 'sequenze'}</span>
            </div>
            <button class="folder-edit-btn" data-id="${f.id}" title="Modifica cartella">✏️</button>
            <button class="folder-delete-btn" data-id="${f.id}" title="Elimina cartella">✕</button>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function renderFoldersEmpty() {
  return `
    <div class="empty-state animate-in">
      <div class="empty-state-icon">📁</div>
      <h3>Nessuna cartella creata</h3>
      <p>Crea una cartella per organizzare le tue sintesi peptidiche</p>
      <button class="btn btn-primary" onclick="document.getElementById('btn-new-folder').click()">
        <span>＋</span> Crea Prima Cartella
      </button>
    </div>
  `;
}

function renderNewFolderModal() {
  const FOLDER_COLORS = ['#06b6d4', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444', '#ec4899', '#3b82f6', '#14b8a6'];
  return `
    <div class="modal-overlay" id="folder-modal-overlay">
      <div class="modal" style="max-width:460px">
        <div class="modal-header">
          <h2>📁 Nuova Cartella</h2>
          <button class="modal-close" id="folder-modal-close">&times;</button>
        </div>
        <div class="modal-body">
          <div class="form-group">
            <label class="form-label">Nome Cartella *</label>
            <input type="text" class="form-input" id="folder-name" placeholder="es. Progetto Alpha, Peptidi Ciclici...">
          </div>
          <div class="form-group">
            <label class="form-label">Colore</label>
            <div class="color-picker">
              ${FOLDER_COLORS.map((c, i) => `
                <label class="color-option">
                  <input type="radio" name="folder-color" value="${c}" ${i === 0 ? 'checked' : ''}>
                  <span class="color-swatch" style="background:${c}"></span>
                </label>
              `).join('')}
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn btn-secondary" id="folder-modal-cancel">Annulla</button>
          <button class="btn btn-primary" id="folder-modal-save">📁 Crea Cartella</button>
        </div>
      </div>
    </div>
  `;
}

// --- ALL FILES TAB ---
function renderAllFilesTab() {
  return `
    <div class="dashboard-toolbar animate-in" style="margin-top: 20px;">
      <h2 class="section-heading">Archivio Globale</h2>
      <button class="btn btn-primary" id="btn-new-synthesis-global"><span>＋</span> Nuova Sintesi (senza cartella)</button>
    </div>

    <div class="search-bar animate-in">
      <span class="search-icon">🔍</span>
      <input type="text" id="search-files-input" placeholder="Cerca globalmente per nome o sequenza..." value="${searchQuery}">
    </div>

    <div id="all-files-container">
      ${renderAllFilesList()}
    </div>
  `;
}

function renderAllFilesList() {
  const filtered = searchQuery 
    ? syntheses.filter(s => s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.sequence.toLowerCase().includes(searchQuery.toLowerCase()))
    : syntheses;

  if (syntheses.length === 0) return `
    <div class="empty-state animate-in">
      <div class="empty-state-icon">🧬</div>
      <h3>Nessuna sintesi creata</h3>
      <p>Aggiungi la tua prima sintesi peptidica.</p>
    </div>`;

  if (filtered.length === 0) return `<div class="empty-state animate-in"><div class="empty-state-icon">🔍</div><h3>Nessun risultato globale</h3></div>`;
  
  return renderTable(filtered, true); 
}

// ============================================================
// FOLDER VIEW — Syntheses inside a folder
// ============================================================
function renderFolderView(folder) {
  const folderSyntheses = syntheses.filter(s => s.folderId === folder.id);
  const filtered = searchQuery
    ? folderSyntheses.filter(s => s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.sequence.toLowerCase().includes(searchQuery.toLowerCase()))
    : folderSyntheses;

  return `
    <div class="dashboard-toolbar animate-in">
      <div>
        <div class="breadcrumb" style="margin-bottom:8px">
          <a id="btn-back-dashboard">← Torna alla Dashboard</a>
          <span>/</span>
          <span>${escapeHtml(folder.name)}</span>
        </div>
        <h2 class="section-heading" style="margin:0">📁 ${escapeHtml(folder.name)}</h2>
      </div>
      <button class="btn btn-primary" id="btn-new-synthesis"><span>＋</span> Nuova Sintesi qui</button>
    </div>

    <div class="search-bar animate-in">
      <span class="search-icon">🔍</span>
      <input type="text" id="search-input" placeholder="Cerca per nome o sequenza in questa cartella..." value="${searchQuery}">
    </div>

    <div id="folder-view-container">
      ${filtered.length === 0 && folderSyntheses.length === 0 ? `
        <div class="empty-state animate-in">
          <div class="empty-state-icon">🧬</div>
          <h3>Cartella vuota</h3>
          <p>Aggiungi la tua prima sintesi in questa cartella</p>
          <button class="btn btn-primary" onclick="document.getElementById('btn-new-synthesis').click()">
            <span>＋</span> Nuova Sintesi
          </button>
        </div>
      ` : filtered.length === 0 ? `
        <div class="empty-state animate-in">
          <div class="empty-state-icon">🔍</div>
          <h3>Nessun risultato in questa cartella</h3>
        </div>
      ` : renderTable(filtered, false)}
    </div>

    ${renderNewSynthesisModal(folder.id)}
  `;
}

// Table Rendering Helper
function renderTable(filtered, showFolder = false) {
  return `
    <table class="synthesis-table animate-in">
      <thead>
        <tr>
          <th>Nome</th>
          ${showFolder ? '<th>Cartella</th>' : ''}
          <th>Sequenza</th>
          <th>Residui</th>
          <th>Data</th>
          <th>Resina</th>
          <th>Stato</th>
        </tr>
      </thead>
      <tbody>
        ${filtered.map(s => {
          const folderObj = showFolder && s.folderId ? folders.find(f=>f.id === s.folderId) : null;
          const folderName = folderObj ? folderObj.name : '—';
          const tokens = tokenizeSequence(s.sequence);
          return `
          <tr data-id="${s.id}">
            <td><strong>${escapeHtml(s.name)}</strong></td>
            ${showFolder ? `<td><span style="font-size:0.8rem; color:var(--text-muted)">${escapeHtml(folderName)}</span></td>` : ''}
            <td><span class="sequence-code">${truncateSequence(s.sequence, 18)}</span></td>
            <td>${tokens.length}</td>
            <td>${formatDate(s.dateStarted)}</td>
            <td>${escapeHtml(s.resinType)}</td>
            <td>${renderStatusBadge(s.status)}</td>
          </tr>
        `;}).join('')}
      </tbody>
    </table>
  `;
}

function renderStatusBadge(status) {
  const map = {
    'in-progress': { label: '⏳ In Corso', cls: 'badge-progress' },
    'completed':   { label: '✅ Completata', cls: 'badge-completed' },
    'cleaved':     { label: '🔬 Cleavage', cls: 'badge-cleaved' },
    'analyzed':    { label: '📊 Analizzata', cls: 'badge-analyzed' }
  };
  const info = map[status] || map['in-progress'];
  return `<span class="badge ${info.cls}">${info.label}</span>`;
}

// ============================================================
// NEW SYNTHESIS MODAL
// ============================================================
function renderNewSynthesisModal(defaultFolderId) {
  return `
    <div class="modal-overlay" id="modal-overlay">
      <div class="modal">
        <div class="modal-header">
          <h2>🧪 Nuova Sintesi Peptidica</h2>
          <button class="modal-close" id="modal-close">&times;</button>
        </div>
        <div class="modal-body">
          <div class="form-group" style="${folders.length === 0 ? 'display:none;' : ''}">
            <label class="form-label">Seleziona Cartella</label>
            <select class="form-select" id="form-folder-id">
              <option value="">— Nessuna Cartella (Archivio Globale) —</option>
              ${folders.map(f => `<option value="${f.id}" ${f.id === defaultFolderId ? 'selected' : ''}>${escapeHtml(f.name)}</option>`).join('')}
            </select>
          </div>

          <div class="form-group">
            <label class="form-label">Nome Sequenza *</label>
            <input type="text" class="form-input" id="form-name" placeholder="es. Peptide-001">
          </div>

          <div class="form-group">
            <label class="form-label">Sequenza Amminoacidica *</label>
            <input type="text" class="form-input mono" id="form-sequence" placeholder="es. KADESFYRWG" autocomplete="off">
            <div class="form-hint">Inserisci la sequenza in codice a una lettera (N→C)</div>
            <div class="form-error" id="seq-error" style="display:none"></div>
          </div>

          <div class="form-group">
            <label class="form-label">Data Inizio</label>
            <input type="date" class="form-input" id="form-date" value="${new Date().toISOString().split('T')[0]}">
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Resina *</label>
              <select class="form-select" id="form-resin">
                ${RESINS.map(r => `<option value="${r.name}" data-type="${r.type}" data-loading="${r.defaultLoading}">${r.name} (${r.range})</option>`).join('')}
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Loading (mmol/g) *</label>
              <input type="number" class="form-input" id="form-loading" step="0.01" value="${RESINS[0].defaultLoading}">
            </div>
          </div>

          <div class="wang-preloaded-panel" id="form-wang-preloaded-group" style="display:none">
            <div class="form-group">
              <label class="form-label">Wang preloaded</label>
              <select class="form-select" id="form-preloaded-residue">
                <option value="">Wang non preloaded / caricamento manuale</option>
                ${renderPreloadedResidueOptions('')}
              </select>
              <div class="form-hint">Se coincide con il residuo C-terminale, non viene calcolato come coupling da pesare.</div>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Scala (mmol) *</label>
              <input type="number" class="form-input" id="form-scale" step="0.01" value="0.05">
            </div>
            <div class="form-group">
              <label class="form-label">Equivalenti AA *</label>
              <input type="number" class="form-input" id="form-eq" step="0.1" value="3">
            </div>
          </div>

          <div class="form-group">
            <label class="form-label" style="margin-bottom:10px">Attivatori (fino a 3) *</label>
            <div class="activators-list">
              <div class="activator-row">
                <span class="activator-badge">1</span>
                <select class="form-select" id="form-activator-1">
                  ${ACTIVATORS.map(a => `<option value="${a.name}">${a.name}</option>`).join('')}
                </select>
                <span class="activator-default-tag">Default</span>
              </div>
              <div class="activator-row">
                <span class="activator-badge">2</span>
                <select class="form-select" id="form-activator-2">
                  <option value="">— Nessuno —</option>
                  ${ACTIVATORS.map(a => `<option value="${a.name}">${a.name}</option>`).join('')}
                </select>
              </div>
              <div class="activator-row">
                <span class="activator-badge">3</span>
                <select class="form-select" id="form-activator-3">
                  <option value="">— Nessuno —</option>
                  ${ACTIVATORS.map(a => `<option value="${a.name}">${a.name}</option>`).join('')}
                </select>
              </div>
            </div>
            <div class="form-hint">L'attivatore 1 è il default. Potrai cambiare l'attivatore per ogni singolo AA nella scheda.</div>
          </div>

          <div id="form-preview" style="display:none"></div>
        </div>
        <div class="modal-footer">
          <button class="btn btn-secondary" id="modal-cancel">Annulla</button>
          <button class="btn btn-primary" id="modal-save">💾 Crea Scheda</button>
        </div>
      </div>
    </div>
  `;
}

// ============================================================
// DETAIL VIEW
// ============================================================
function renderDetail(s) {
  const resin = RESINS.find(r => r.name === s.resinType);
  const cTerminus = resin ? resin.type : 'amide';
  const tokens = tokenizeSequence(s.sequence);
  const cyclizationInfo = analyzeCyclizations(s.sequence);
  const cyclizationProtectedOverrides = getCyclizationProtectedOverrides(s.sequence);
  const peptideMW = calculatePeptideMW(tokens, cTerminus, s.customMWs, s.sequence);
  const resinMass = calculateResinMass(s.scale, s.resinLoading);
  const cleavageInfo = suggestCleavageCocktail(tokens);
  const adducts = calculateMSAdducts(peptideMW);
  const availableActivators = getAvailableActivators(s);
  if (!s.aaActivators) s.aaActivators = {};

  const folder = folders.find(f => f.id === s.folderId);
  const seqReversed = [...tokens].reverse();
  
  if (!s.customMWs) s.customMWs = {};
  let isFetching = false;
  let couplingRows = '';
  
  seqReversed.forEach((aa, idx) => {
    const forwardIdx = tokens.length - 1 - idx;
    const protectedOverride = cyclizationProtectedOverrides[forwardIdx];
    const isPreloadedStep = isPreloadedWangStep(s, aa, forwardIdx, tokens.length);
    let aaData = AMINO_ACIDS[aa];
    let fmocMW = 0;
    let fallbackInput = '';

    if (!aaData) {
      // Unconventional AA
      const inv = findInInventory(aa);
      const cas = inv ? inv.cas : aa;
      
      fmocMW = parseFloat(s.customMWs[aa]);
      if (isNaN(fmocMW)) {
        fmocMW = 0;
        // Trigger PubChem async fetch if not already tried
        if (s.customMWs[aa] !== 'failed' && s.customMWs[aa] !== 'fetching') {
          s.customMWs[aa] = 'fetching';
          fetchPubChemMW(cas, aa, s.id);
          isFetching = true;
        }
      }

      if (s.customMWs[aa] === 'failed' || s.customMWs[aa] === 0) {
         fallbackInput = `<input type="number" class="form-input" style="width:70px; padding:2px; font-size:0.8rem; display:inline;" placeholder="MW" value="" onchange="saveCustomMW('${s.id}', '${aa}', this.value)">`;
      }

      aaData = {
        code3: 'Spez.',
        fmocName: aa,
        fmocMW: fmocMW
      };
    } else {
      fmocMW = aaData.fmocMW;
    }

    if (aaData && protectedOverride) {
      aaData = Object.assign({}, aaData, {
        code3: protectedOverride.code3 || aaData.code3,
        fmocName: protectedOverride.fmocName,
        fmocMW: protectedOverride.fmocMW,
        protection: protectedOverride.protection
      });
      fmocMW = aaData.fmocMW;
    }

    const aaMass = calculateAAMass(aa, s.equivalents, s.scale, fmocMW);
    const aaActName = s.aaActivators[idx] || availableActivators[0]?.name || '';
    const aaActivator = ACTIVATORS.find(a => a.name === aaActName) || availableActivators[0] || ACTIVATORS[0];
    const actMass = calculateActivatorMass(aaActivator, s.scale);
    const actOptions = availableActivators.map(a =>
      `<option value="${a.name}" ${a.name === aaActName ? 'selected' : ''}>${a.name}</option>`
    ).join('');
    const activatorCell = isPreloadedStep ? '<span class="preloaded-tag">Precaricato</span>' : `<select class="row-activator-select" data-idx="${idx}">${actOptions}</select>`;
    const eqCell = isPreloadedStep ? '-' : s.equivalents;
    const aaMassCell = isPreloadedStep ? '0.00' : (s.customMWs && s.customMWs[aa] === 'fetching' ? '⏳' : aaMass.toFixed(2));
    const actMassCell = isPreloadedStep ? '-' : actMass.activatorMass.toFixed(2);

    couplingRows += `
      <tr>
        <td>${idx + 1}</td>
        <td><strong>${aaData.code3}</strong> (${aa})</td>
        <td>${aaData.fmocName}</td>
        <td class="mw-value">${s.customMWs && s.customMWs[aa] === 'fetching' ? '⏳ <i>PubChem...</i>' : (fallbackInput ? fallbackInput : aaData.fmocMW.toFixed(2))}</td>
        <td>${eqCell}</td>
        <td class="mass-value">${aaMassCell}</td>
        <td>${activatorCell}</td>
        <td class="mass-value">${actMassCell}</td>
      </tr>`;
  });

  const inventoryTokens = tokens.map((token, idx) => {
    return cyclizationProtectedOverrides[idx] ? cyclizationProtectedOverrides[idx].fmocName : token;
  });
  const uniqueTokens = [...new Set(inventoryTokens)];
  const usedActivatorNames = new Set();
  seqReversed.forEach((aa, idx) => {
    const forwardIdx = tokens.length - 1 - idx;
    if (!isPreloadedWangStep(s, aa, forwardIdx, tokens.length)) {
      usedActivatorNames.add(s.aaActivators[idx] || availableActivators[0]?.name || '');
    }
  });
  
  let reagentRows = '';
  usedActivatorNames.forEach(actName => {
    const act = ACTIVATORS.find(a => a.name === actName);
    if (!act) return;
    if (act.base) reagentRows += `<tr class="resin-row"><td colspan="2"></td><td><strong>${act.base}</strong> (base per ${act.name})</td><td class="mw-value">${act.baseMW}</td><td>${act.baseEq}</td><td class="mass-value">${(act.baseMW * act.baseEq * s.scale).toFixed(2)} mg</td><td colspan="2" style="font-size:0.8rem;color:var(--text-muted)">≈ ${((act.baseMW * act.baseEq * s.scale) / 0.742).toFixed(1)} µL per coupling</td></tr>`;
    if (act.coReagent) reagentRows += `<tr class="resin-row"><td colspan="2"></td><td><strong>${act.coReagent}</strong> (per ${act.name})</td><td class="mw-value">${act.coReagentMW}</td><td>${act.coReagentEq}</td><td class="mass-value">${(act.coReagentMW * act.coReagentEq * s.scale).toFixed(2)} mg</td><td colspan="2"></td></tr>`;
  });

  const synthFiles = attachments[s.id] || { ms: [], hplc: [] };

  return `
    <div class="detail-header animate-in">
      <div class="detail-header-left">
        <div class="breadcrumb">
          <a id="btn-back-dashboard">Dashboard</a>
          ${folder ? `<span>/</span><a id="btn-back-folder">${escapeHtml(folder.name)}</a>` : '<span>/</span><a id="btn-back-all">Tutte le Sequenze</a>'}
          <span>/</span>
          <span>${escapeHtml(s.name)}</span>
        </div>
        <h2 class="detail-title">${escapeHtml(s.name)}</h2>
        <div class="detail-sequence-container">
          <div class="detail-sequence">${renderSequenceWithCyclization(s.sequence)}</div>
          <button class="btn btn-secondary btn-sm" id="btn-edit-sequence" title="Modifica sequenza">✏️ Modifica</button>
        </div>
      </div>
      <div class="detail-header-actions">
        <select class="form-select btn-sm" id="status-select" style="width:auto; padding:6px 30px 6px 12px;">
          <option value="in-progress" ${s.status === 'in-progress' ? 'selected' : ''}>⏳ In Corso</option>
          <option value="completed" ${s.status === 'completed' ? 'selected' : ''}>✅ Completata</option>
          <option value="cleaved" ${s.status === 'cleaved' ? 'selected' : ''}>🔬 Cleavage</option>
          <option value="analyzed" ${s.status === 'analyzed' ? 'selected' : ''}>📊 Analizzata</option>
        </select>
        <button class="btn btn-secondary btn-sm" id="btn-edit-params">Modifica Parametri</button>
        <button class="btn btn-secondary btn-sm" id="btn-print-lab">🖨️ Stampa Foglio Lab</button>
        <button class="btn btn-danger btn-sm" id="btn-delete">🗑️</button>
      </div>
    </div>

    <!-- Stats & Info -->
    <div class="info-grid animate-in">
      <div class="info-item"><div class="info-item-label">Data Inizio</div><div class="info-item-value">${formatDate(s.dateStarted)}</div></div>
      <div class="info-item"><div class="info-item-label">Residui</div><div class="info-item-value">${tokens.length} AA</div></div>
      <div class="info-item"><div class="info-item-label">Resina</div><div class="info-item-value">${escapeHtml(s.resinType)}</div></div>
      ${isWangResinName(s.resinType) && s.preloadedResidue ? `<div class="info-item"><div class="info-item-label">Wang Preloaded</div><div class="info-item-value">${escapeHtml(getPreloadedResidueLabel(s.preloadedResidue))}</div></div>` : ''}
      <div class="info-item"><div class="info-item-label">Loading</div><div class="info-item-value mono">${s.resinLoading} mmol/g</div></div>
      <div class="info-item"><div class="info-item-label">Massa Resina</div><div class="info-item-value mono">${resinMass.toFixed(1)} mg</div></div>
      <div class="info-item"><div class="info-item-label">Scala</div><div class="info-item-value mono">${s.scale} mmol</div></div>
      <div class="info-item"><div class="info-item-label">Equivalenti</div><div class="info-item-value mono">${s.equivalents} eq</div></div>
      <div class="info-item"><div class="info-item-label">Attivatori</div><div class="info-item-value">${availableActivators.map(a => a.name).join(', ')}</div></div>
      ${cyclizationInfo.cyclizations.length > 0 ? `<div class="info-item"><div class="info-item-label">Ciclizzazione</div><div class="info-item-value">${cyclizationInfo.cyclizations.length} ponte${cyclizationInfo.cyclizations.length > 1 ? 'i' : ''}</div></div>` : ''}
    </div>

    ${renderPreloadedResinWarning(s, tokens)}

    <!-- Calc Table -->
    <div class="section-card animate-in">
      <div class="section-card-header"><div class="section-card-title">📊 Tabella di Calcolo — Coupling Steps</div><span style="font-size:0.78rem;color:var(--text-muted)">Ordine: C→N (SPPS)</span></div>
      <div class="section-card-body" style="padding:0;overflow-x:auto">
        <table class="calc-table"><thead><tr><th>#</th><th>AA</th><th>Fmoc-AA</th><th>PM (g/mol)</th><th>Eq.</th><th>Massa AA (mg)</th><th>Attivatore</th><th>Massa Att. (mg)</th></tr></thead>
        <tbody>${couplingRows}${reagentRows}</tbody></table>
      </div>
    </div>

    <!-- INVENTORY CHECK -->
    <div class="section-card animate-in">
      <div class="section-card-header"><div class="section-card-title">🔍 Verifica Magazzino Inventario</div><span style="font-size:0.78rem;color:var(--text-muted)">In base file Excel collegato</span></div>
      <div class="section-card-body" style="padding:0;overflow-x:auto">
        <table class="calc-table">
          <thead><tr><th>Composto</th><th>Quantità (Conf.)</th><th>Marca</th><th>Codice</th><th>Locazione</th><th>CAS</th><th>Residuo</th></tr></thead>
          <tbody>
            ${renderInventoryCheckRows(uniqueTokens, Array.from(usedActivatorNames), s)}
          </tbody>
        </table>
        ${(!inventoryData || inventoryData.length === 0) ? `<div style="padding:16px; text-align:center; color:var(--text-muted);">Nessun inventario Excel collegato. Usa il pulsante nella dashboard.</div>` : ''}
      </div>
    </div>

    <!-- MS Section -->
    <div class="section-card animate-in">
      <div class="section-card-header"><div class="section-card-title">⚗️ Peso Molecolare Peptide & Spettrometria di Massa</div></div>
      <div class="section-card-body">
        <div class="info-grid" style="margin-bottom:20px">
          <div class="info-item"><div class="info-item-label">PM Peptide Atteso</div><div class="info-item-value mono" style="color:var(--accent);font-size:1.3rem">${peptideMW.toFixed(2)} Da</div></div>
          <div class="info-item"><div class="info-item-label">C-Terminale</div><div class="info-item-value">${cTerminus === 'amide' ? 'Ammide (-NH₂)' : 'Acido (-OH)'}</div></div>
          <div class="info-item"><div class="info-item-label">Formula</div><div class="info-item-value mono" style="font-size:0.85rem">H-${escapeHtml(cyclizationInfo.cleanSequence)}-${cTerminus === 'amide' ? 'NH₂' : 'OH'}</div></div>
          ${cyclizationInfo.cyclizations.length > 0 ? `<div class="info-item"><div class="info-item-label">Delta Ciclizzazione</div><div class="info-item-value mono">${formatMassDelta(cyclizationInfo.totalMassDelta)} Da</div></div>` : ''}
        </div>
        <h4 style="font-size:0.82rem;color:var(--text-muted);text-transform:uppercase;margin-bottom:12px">Addotti MS attesi</h4>
        <div class="ms-grid">
          ${adducts.map(a => `<div class="ms-adduct"><div class="ms-adduct-name">${a.name}</div><div class="ms-adduct-value">${a.mz.toFixed(2)}</div></div>`).join('')}
        </div>
        <div style="margin-top:16px">
          <label class="form-label">MS Osservato (m/z)</label>
          <input type="text" class="form-input mono" id="ms-observed" placeholder="m/z osservato" value="${s.msObserved || ''}" style="max-width:300px">
          <div class="form-hint" id="ms-match"></div>
        </div>
        <div class="attachment-section">
          <div class="attachment-header"><h4>📎 Spettri MS Allegati</h4>
            <label class="btn btn-secondary btn-sm attachment-upload-btn"><span>＋</span> Aggiungi Spettro<input type="file" id="ms-file-input" accept="image/*,.pdf" multiple style="display:none"></label>
          </div>
          <div class="attachment-grid" id="ms-attachments">${renderAttachments(synthFiles.ms, 'ms')}</div>
        </div>
      </div>
    </div>

    ${renderCyclizationSection(cyclizationInfo, s)}

    <!-- Cleavage & Notes -->
    <div class="section-card animate-in">
      <div class="section-card-header"><div class="section-card-title">🧪 Cleavage — Suggerimento</div></div>
      <div class="section-card-body">
        <div class="cleavage-card">
          <h4>💡 ${cleavageInfo.cocktail.name}</h4>
          <div class="cleavage-detail"><span><strong>Cocktail:</strong> ${cleavageInfo.cocktail.composition}</span><span><strong>Tempo:</strong> ${cleavageInfo.cocktail.time}</span></div>
          <p style="font-size:0.85rem;color:var(--text-secondary)">${cleavageInfo.cocktail.notes}</p>
          ${cleavageInfo.warnings.length > 0 ? `<ul class="cleavage-warnings">${cleavageInfo.warnings.map(w => `<li>${w}</li>`).join('')}</ul>` : ''}
        </div>
      </div>
    </div>
    <div class="section-card animate-in">
      <div class="section-card-header"><div class="section-card-title">📝 Note & Procedure</div></div>
      <div class="section-card-body">
        <textarea class="notes-textarea" id="notes-textarea" placeholder="Aggiungi note...">${escapeHtml(s.notes || '')}</textarea>
        <div class="notes-saved" id="notes-saved">✓ Note salvate</div>
      </div>
    </div>

    <!-- HPLC -->
    <div class="section-card animate-in">
      <div class="section-card-header"><div class="section-card-title">📈 HPLC Analitico</div></div>
      <div class="section-card-body">
        <div class="form-row" style="max-width:500px">
          <div class="form-group"><label class="form-label">Purezza (%)</label><input type="number" class="form-input" id="hplc-purity" step="0.1" value="${s.hplcPurity || ''}"></div>
          <div class="form-group"><label class="form-label">Tempo Ritenzione (min)</label><input type="number" class="form-input" id="hplc-rt" step="0.01" value="${s.hplcRT || ''}"></div>
        </div>
        <div class="form-group" style="max-width:500px">
          <label class="form-label">Note HPLC</label><textarea class="notes-textarea" id="hplc-notes" style="min-height:80px">${escapeHtml(s.hplcNotes || '')}</textarea>
        </div>
        <div class="attachment-section">
          <div class="attachment-header"><h4>📎 Cromatogrammi HPLC Allegati</h4>
            <label class="btn btn-secondary btn-sm attachment-upload-btn"><span>＋</span> Aggiungi Cromatogramma<input type="file" id="hplc-file-input" accept="image/*,.pdf" multiple style="display:none"></label>
          </div>
          <div class="attachment-grid" id="hplc-attachments">${renderAttachments(synthFiles.hplc, 'hplc')}</div>
        </div>
      </div>
    </div>
    ${renderEditSynthesisModal(s)}
  `;
}

function renderEditSynthesisModal(s) {
  const activeActivators = getAvailableActivators(s);
  const a1 = activeActivators[0]?.name || ACTIVATORS[0].name;
  const a2 = activeActivators[1]?.name || '';
  const a3 = activeActivators[2]?.name || '';

  return `
    <div class="modal-overlay" id="edit-modal-overlay">
      <div class="modal">
        <div class="modal-header">
          <h2>Modifica Parametri Sintesi</h2>
          <button class="modal-close" id="edit-modal-close">&times;</button>
        </div>
        <div class="modal-body">
          <div class="form-group" style="${folders.length === 0 ? 'display:none;' : ''}">
            <label class="form-label">Cartella</label>
            <select class="form-select" id="edit-folder-id">
              <option value="">Archivio Globale</option>
              ${folders.map(f => `<option value="${f.id}" ${f.id === s.folderId ? 'selected' : ''}>${escapeHtml(f.name)}</option>`).join('')}
            </select>
          </div>

          <div class="form-group">
            <label class="form-label">Nome Sequenza *</label>
            <input type="text" class="form-input" id="edit-name" value="${escapeHtml(s.name)}">
          </div>

          <div class="form-group">
            <label class="form-label">Sequenza Amminoacidica *</label>
            <input type="text" class="form-input mono" id="edit-sequence" value="${escapeHtml(s.sequence)}" autocomplete="off">
            <div class="form-error" id="edit-seq-error" style="display:none"></div>
          </div>

          <div class="form-group">
            <label class="form-label">Data Inizio</label>
            <input type="date" class="form-input" id="edit-date" value="${escapeHtml(s.dateStarted || '')}">
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Resina *</label>
              <select class="form-select" id="edit-resin">
                ${RESINS.map(r => `<option value="${r.name}" data-type="${r.type}" data-loading="${r.defaultLoading}" ${r.name === s.resinType ? 'selected' : ''}>${r.name} (${r.range})</option>`).join('')}
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Loading (mmol/g) *</label>
              <input type="number" class="form-input" id="edit-loading" step="0.01" value="${s.resinLoading}">
            </div>
          </div>

          <div class="wang-preloaded-panel" id="edit-wang-preloaded-group" style="display:none">
            <div class="form-group">
              <label class="form-label">Wang preloaded</label>
              <select class="form-select" id="edit-preloaded-residue">
                <option value="">Wang non preloaded / caricamento manuale</option>
                ${renderPreloadedResidueOptions(s.preloadedResidue || '')}
              </select>
              <div class="form-hint">Se coincide con il residuo C-terminale, non viene calcolato come coupling da pesare.</div>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Scala (mmol) *</label>
              <input type="number" class="form-input" id="edit-scale" step="0.01" value="${s.scale}">
            </div>
            <div class="form-group">
              <label class="form-label">Equivalenti AA *</label>
              <input type="number" class="form-input" id="edit-eq" step="0.1" value="${s.equivalents}">
            </div>
          </div>

          <div class="form-group">
            <label class="form-label" style="margin-bottom:10px">Attivatori (fino a 3) *</label>
            <div class="activators-list">
              <div class="activator-row">
                <span class="activator-badge">1</span>
                <select class="form-select" id="edit-activator-1">
                  ${ACTIVATORS.map(a => `<option value="${a.name}" ${a.name === a1 ? 'selected' : ''}>${a.name}</option>`).join('')}
                </select>
                <span class="activator-default-tag">Default</span>
              </div>
              <div class="activator-row">
                <span class="activator-badge">2</span>
                <select class="form-select" id="edit-activator-2">
                  <option value="">Nessuno</option>
                  ${ACTIVATORS.map(a => `<option value="${a.name}" ${a.name === a2 ? 'selected' : ''}>${a.name}</option>`).join('')}
                </select>
              </div>
              <div class="activator-row">
                <span class="activator-badge">3</span>
                <select class="form-select" id="edit-activator-3">
                  <option value="">Nessuno</option>
                  ${ACTIVATORS.map(a => `<option value="${a.name}" ${a.name === a3 ? 'selected' : ''}>${a.name}</option>`).join('')}
                </select>
              </div>
            </div>
          </div>

          <div id="edit-preview" style="display:none"></div>
        </div>
        <div class="modal-footer">
          <button class="btn btn-secondary" id="edit-modal-cancel">Annulla</button>
          <button class="btn btn-primary" id="edit-modal-save">Salva Modifiche</button>
        </div>
      </div>
    </div>
  `;
}

// ============================================================
// ATTACHMENTS LOGIC
// ============================================================
function renderAttachments(files, type) {
  if (!files || files.length === 0) return '';
  return files.map((file, idx) => `
    <div class="attachment-card" data-type="${type}" data-idx="${idx}">
      <div class="attachment-preview">
        ${file.dataUrl.startsWith('data:image/') ? `<img src="${file.dataUrl}" alt="${escapeHtml(file.name)}" class="attachment-img" onclick="openImageModal(this.src, '${escapeHtml(file.name)}')">` : `<div class="attachment-file-icon">📄</div>`}
      </div>
      <div class="attachment-info"><span class="attachment-name" title="${escapeHtml(file.name)}">${truncateText(file.name, 25)}</span><span class="attachment-size">${formatFileSize(file.size)}</span></div>
      <button class="attachment-delete" data-type="${type}" data-idx="${idx}" title="Rimuovi">✕</button>
    </div>
  `).join('');
}
function openImageModal(src, name) {
  const overlay = document.createElement('div'); overlay.className = 'image-modal-overlay';
  overlay.innerHTML = `<div class="image-modal"><div class="image-modal-header"><span>${name}</span><button class="modal-close" onclick="this.closest('.image-modal-overlay').remove()">&times;</button></div><img src="${src}" alt="${name}" style="max-width:100%;max-height:80vh;border-radius:8px;"></div>`;
  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); }); document.body.appendChild(overlay);
}
async function handleFileUpload(files, synthesisId, type) {
  if (!attachments[synthesisId]) attachments[synthesisId] = { ms: [], hplc: [] };
  for (const file of files) {
    if (file.size > 5*1024*1024) continue;
    const dataUrl = await readFileAsDataUrl(file);
    attachments[synthesisId][type].push({ name: file.name, size: file.size, mimeType: file.type, dataUrl, addedAt: new Date().toISOString() });
  }
  saveFiles(attachments); render();
}
function readFileAsDataUrl(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file); }); }
function deleteAttachment(synthesisId, type, idx) { if (!attachments[synthesisId]) return; attachments[synthesisId][type].splice(idx, 1); saveFiles(attachments); render(); }

// ============================================================
// EVENT BINDINGS
// ============================================================

function bindDashboardEvents() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      currentTab = e.target.dataset.tab;
      searchQuery = '';
      render();
    });
  });

  const btnNewFolder = document.getElementById('btn-new-folder');
  if (btnNewFolder) btnNewFolder.addEventListener('click', () => { document.getElementById('folder-modal-overlay').classList.add('active'); document.getElementById('folder-name').focus(); });

  const btnNewSynthGlobal = document.getElementById('btn-new-synthesis-global');
  if (btnNewSynthGlobal) btnNewSynthGlobal.addEventListener('click', () => { document.getElementById('modal-overlay').classList.add('active'); document.getElementById('form-folder-id').value = ''; });

  const dFolderClose = document.getElementById('folder-modal-close');
  if (dFolderClose) {
    dFolderClose.addEventListener('click', () => closeFolderModal());
    document.getElementById('folder-modal-cancel').addEventListener('click', () => closeFolderModal());
    document.getElementById('folder-modal-overlay').addEventListener('click', e => { if (e.target.id === 'folder-modal-overlay') closeFolderModal(); });
    
    document.getElementById('folder-modal-save').addEventListener('click', () => {
      const name = document.getElementById('folder-name').value.trim();
      if (!name) return;
      const color = document.querySelector('input[name="folder-color"]:checked')?.value || '#06b6d4';
      const existingId = document.getElementById('folder-modal-overlay').dataset.editId;
      
      if (existingId) {
        const f = folders.find(x => x.id === existingId);
        if (f) { f.name = name; f.color = color; }
      } else {
        folders.push({ id: generateId(), name, color, createdAt: new Date().toISOString() });
      }
      
      saveFolders(folders);
      closeFolderModal();
      render();
    });
  }

  function closeFolderModal() {
    const overlay = document.getElementById('folder-modal-overlay');
    overlay.classList.remove('active');
    delete overlay.dataset.editId;
    overlay.querySelector('h2').innerText = '📁 Nuova Cartella';
    overlay.querySelector('#folder-modal-save').innerText = '📁 Crea Cartella';
    document.getElementById('folder-name').value = '';
  }
  
  window.openEditFolderModal = (folder) => {
    const overlay = document.getElementById('folder-modal-overlay');
    overlay.dataset.editId = folder.id;
    overlay.querySelector('h2').innerText = '✏️ Modifica Cartella';
    overlay.querySelector('#folder-modal-save').innerText = '📝 Salva Modifiche';
    document.getElementById('folder-name').value = folder.name;
    const radio = overlay.querySelector(`input[name="folder-color"][value="${folder.color}"]`);
    if (radio) radio.checked = true;
    overlay.classList.add('active');
  };

  const sFolderInput = document.getElementById('search-folder-input');
  if (sFolderInput) {
    sFolderInput.addEventListener('input', e => {
      searchQuery = e.target.value;
      const container = document.getElementById('folders-container');
      if (container) {
        container.innerHTML = renderFoldersGridList();
        bindFolderCards();
      }
    });
  }

  const sFilesInput = document.getElementById('search-files-input');
  if (sFilesInput) {
    sFilesInput.addEventListener('input', e => {
      searchQuery = e.target.value;
      const container = document.getElementById('all-files-container');
      if (container) {
        container.innerHTML = renderAllFilesList();
        bindTableRowClicks();
      }
    });
  }

  bindFolderCards();
  bindTableRowClicks();
  bindSynthesisModalEvents();
}

function bindFolderCards() {
  document.querySelectorAll('.folder-card').forEach(card => {
    card.addEventListener('click', e => {
      if (e.target.closest('.folder-delete-btn') || e.target.closest('.folder-edit-btn')) return;
      navigate('folder', card.dataset.id);
    });
  });
  
  document.querySelectorAll('.folder-edit-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const folder = folders.find(f => f.id === btn.dataset.id);
      if (folder) window.openEditFolderModal(folder);
    });
  });

  document.querySelectorAll('.folder-delete-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const fId = btn.dataset.id;
      const count = syntheses.filter(s => s.folderId === fId).length;
      if (count > 0 && !confirm(`Eliminare la cartella e TUTTE le ${count} sintesi interne?`)) return;
      if (count === 0 && !confirm(`Eliminare la cartella?`)) return;
      syntheses = syntheses.filter(s => s.folderId !== fId);
      folders = folders.filter(f => f.id !== fId);
      saveSyntheses(syntheses); saveFolders(folders); render();
    });
  });
}

function bindFolderViewEvents(folder) {
  document.getElementById('btn-back-dashboard').addEventListener('click', () => navigate('dashboard'));
  document.getElementById('btn-new-synthesis').addEventListener('click', () => {
    document.getElementById('modal-overlay').classList.add('active');
    document.getElementById('form-folder-id').value = folder.id;
  });

  const searchInput = document.getElementById('search-input');
  if (searchInput) {
    searchInput.addEventListener('input', e => {
      searchQuery = e.target.value;
      const folderSyntheses = syntheses.filter(s => s.folderId === folder.id);
      const filtered = searchQuery ? folderSyntheses.filter(s => s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.sequence.toLowerCase().includes(searchQuery.toLowerCase())) : folderSyntheses;
      const container = document.getElementById('folder-view-container');
      if (container) {
        container.innerHTML = (filtered.length === 0 && folderSyntheses.length > 0) ? `<div class="empty-state animate-in"><h3>Nessun risultato</h3></div>` : renderTable(filtered, false);
        bindTableRowClicks();
      }
    });
  }

  bindTableRowClicks();
  bindSynthesisModalEvents();
}

function bindDetailEvents(synthesis) {
  document.getElementById('btn-back-dashboard').addEventListener('click', () => navigate('dashboard'));
  const btnBackFolder = document.getElementById('btn-back-folder');
  if (btnBackFolder) btnBackFolder.addEventListener('click', () => navigate('folder', synthesis.folderId));
  const btnBackAll = document.getElementById('btn-back-all');
  if (btnBackAll) btnBackAll.addEventListener('click', () => { currentTab = 'files'; navigate('dashboard'); });

  document.getElementById('status-select').addEventListener('change', e => { synthesis.status = e.target.value; saveSyntheses(syntheses); });
  
  const btnEditSeq = document.getElementById('btn-edit-sequence');
  if (btnEditSeq) {
    btnEditSeq.addEventListener('click', () => {
      const newSeq = prompt("Modifica la sequenza peptidica (es. KADESFYRWG, K(Fmoc-O2Oc-OH)G o KR*CVQRC*KDFLR):", synthesis.sequence);
      if (newSeq !== null && newSeq.trim() !== "" && newSeq !== synthesis.sequence) {
        const cycleInfo = analyzeCyclizations(newSeq.trim());
        if (!cycleInfo.isValid) {
          alert(cycleInfo.warnings[0] || 'Controlla gli asterischi della ciclizzazione.');
          return;
        }
        synthesis.sequence = newSeq.trim();
        saveSyntheses(syntheses);
        render();
      }
    });
  }
  const btnEditParams = document.getElementById('btn-edit-params');
  if (btnEditParams) btnEditParams.addEventListener('click', () => {
    document.getElementById('edit-modal-overlay').classList.add('active');
    updateWangPreloadedVisibility('edit');
    updateEditPreview();
  });
  bindEditSynthesisModalEvents(synthesis);

  const btnPrintLab = document.getElementById('btn-print-lab');
  if (btnPrintLab) btnPrintLab.addEventListener('click', () => printSynthesis(synthesis.id));
  document.getElementById('btn-delete').addEventListener('click', () => {
    if (confirm(`Eliminare la sintesi?`)) {
      const fId = synthesis.folderId;
      syntheses = syntheses.filter(s => s.id !== synthesis.id); delete attachments[synthesis.id];
      saveFiles(attachments); saveSyntheses(syntheses);
      if (fId) navigate('folder', fId); else navigate('dashboard');
    }
  });

  document.querySelectorAll('.row-activator-select').forEach(sel => sel.addEventListener('change', e => {
    if (!synthesis.aaActivators) synthesis.aaActivators={};
    synthesis.aaActivators[e.target.dataset.idx] = e.target.value; saveSyntheses(syntheses); render();
  }));

  let tNote; document.getElementById('notes-textarea').addEventListener('input', e => { clearTimeout(tNote); tNote=setTimeout(()=>{synthesis.notes=e.target.value;saveSyntheses(syntheses);showSaved('notes-saved');},600);});
  document.getElementById('ms-observed').addEventListener('input', e => { synthesis.msObserved=e.target.value;saveSyntheses(syntheses);checkMSMatch(synthesis);});
  checkMSMatch(synthesis);

  const saveHPLC = () => { synthesis.hplcPurity=document.getElementById('hplc-purity')?.value; synthesis.hplcRT=document.getElementById('hplc-rt')?.value; synthesis.hplcNotes=document.getElementById('hplc-notes')?.value; saveSyntheses(syntheses); };
  ['hplc-purity','hplc-rt','hplc-notes'].forEach(id=>{const el=document.getElementById(id); if(el){let t;el.addEventListener('input',()=>{clearTimeout(t);t=setTimeout(saveHPLC,600);});}});

  const msI=document.getElementById('ms-file-input'); if(msI) msI.addEventListener('change',e=>{if(e.target.files.length) handleFileUpload(e.target.files,synthesis.id,'ms');});
  const hplcI=document.getElementById('hplc-file-input'); if(hplcI) hplcI.addEventListener('change',e=>{if(e.target.files.length) handleFileUpload(e.target.files,synthesis.id,'hplc');});
  document.querySelectorAll('.attachment-delete').forEach(btn=>btn.addEventListener('click',e=>{e.stopPropagation();if(confirm('Rimuovere?')) deleteAttachment(synthesis.id,btn.dataset.type,parseInt(btn.dataset.idx));}));
}

function bindSynthesisModalEvents() {
  const overlay = document.getElementById('modal-overlay');
  if (!overlay) return;
  const cMod = () => overlay.classList.remove('active');
  document.getElementById('modal-close').addEventListener('click', cMod);
  document.getElementById('modal-cancel').addEventListener('click', cMod);
  overlay.addEventListener('click', e => { if (e.target.id === 'modal-overlay') cMod(); });

  document.getElementById('form-sequence').addEventListener('input', () => { 
    validateSequence(); 
    updatePreview(); 
  });
  document.getElementById('form-resin').addEventListener('change', e => { 
    const sel = e.target.selectedOptions[0];
    if (sel) document.getElementById('form-loading').value = sel.dataset.loading;
    updateWangPreloadedVisibility('form');
    updatePreview();
  });
  ['form-scale','form-eq','form-loading','form-preloaded-residue','form-activator-1','form-activator-2','form-activator-3'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', updatePreview);
      el.addEventListener('change', updatePreview);
    }
  });
  updateWangPreloadedVisibility('form');
  document.getElementById('modal-save').addEventListener('click', saveSynthesis);
}

function bindEditSynthesisModalEvents(synthesis) {
  const overlay = document.getElementById('edit-modal-overlay');
  if (!overlay) return;
  const closeModal = () => overlay.classList.remove('active');
  document.getElementById('edit-modal-close').addEventListener('click', closeModal);
  document.getElementById('edit-modal-cancel').addEventListener('click', closeModal);
  overlay.addEventListener('click', e => { if (e.target.id === 'edit-modal-overlay') closeModal(); });

  document.getElementById('edit-sequence').addEventListener('input', () => {
    validateSequenceField('edit');
    updateEditPreview();
  });
  document.getElementById('edit-resin').addEventListener('change', e => {
    const sel = e.target.selectedOptions[0];
    if (sel) document.getElementById('edit-loading').value = sel.dataset.loading;
    updateWangPreloadedVisibility('edit');
    updateEditPreview();
  });
  ['edit-scale','edit-eq','edit-loading','edit-preloaded-residue','edit-activator-1','edit-activator-2','edit-activator-3'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', updateEditPreview);
      el.addEventListener('change', updateEditPreview);
    }
  });
  document.getElementById('edit-modal-save').addEventListener('click', () => saveEditedSynthesis(synthesis.id));
  updateWangPreloadedVisibility('edit');
}

function bindTableRowClicks() { 
  document.querySelectorAll('.synthesis-table tbody tr').forEach(row => {
    row.addEventListener('click', () => navigate('detail', row.dataset.id));
  }); 
}

// ============================================================
// FORM LOGIC & HELPERS
// ============================================================

function validateSequence() {
  return validateSequenceField('form');
}

function validateSequenceField(prefix) {
  const sq = document.getElementById(prefix + '-sequence').value;
  const er = document.getElementById(prefix === 'form' ? 'seq-error' : prefix + '-seq-error');
  if (!sq) { er.style.display = 'none'; return true; }
  const cyclizationInfo = analyzeCyclizations(sq);
  if (!cyclizationInfo.isValid) {
    er.textContent = cyclizationInfo.warnings[0] || 'Controlla gli asterischi della ciclizzazione';
    er.style.display = 'block';
    return false;
  }
  const tokens = tokenizeSequence(sq);
  for (const token of tokens) {
    if (token.length === 1 && !AMINO_ACIDS[token]) {
      er.textContent = `AA "${token}" non riconosciuto`;
      er.style.display = 'block';
      return false;
    }
  }
  er.style.display = 'none';
  return true;
}

function updatePreview() {
  updateParameterPreview('form');
}

function updateEditPreview() {
  updateParameterPreview('edit');
}

function updateParameterPreview(prefix) {
  const sq = document.getElementById(prefix + '-sequence').value;
  const sc = parseFloat(document.getElementById(prefix + '-scale').value) || 0;
  const eq = parseFloat(document.getElementById(prefix + '-eq').value) || 0;
  const ld = parseFloat(document.getElementById(prefix + '-loading').value) || 0;
  const rt = document.getElementById(prefix + '-resin').selectedOptions[0]?.dataset.type || 'amide';
  const resinName = document.getElementById(prefix + '-resin').value;
  const preloadedResidue = isWangResinName(resinName) ? (document.getElementById(prefix + '-preloaded-residue')?.value || '') : '';
  const pr = document.getElementById(prefix + '-preview');
  if (!sq || !sc || !eq || !ld) { pr.style.display = 'none'; return; }
  const tokens = tokenizeSequence(sq);
  const cyclizationInfo = analyzeCyclizations(sq);
  const pmw = calculatePeptideMW(tokens, rt, {}, sq);
  const rm = calculateResinMass(sc, ld);
  const cycleRow = cyclizationInfo.cyclizations.length > 0
    ? `<div class="form-preview-row"><span>Ciclizzazione:</span><span>${formatMassDelta(cyclizationInfo.totalMassDelta)} Da</span></div>`
    : '';
  const preloadedRow = preloadedResidue
    ? `<div class="form-preview-row"><span>Wang preloaded:</span><span>${escapeHtml(getPreloadedResidueLabel(preloadedResidue))}</span></div>${getPreloadedMismatchMessage(preloadedResidue, tokens) ? `<div class="form-preview-warning">${escapeHtml(getPreloadedMismatchMessage(preloadedResidue, tokens))}</div>` : ''}`
    : '';
  pr.style.display = 'block';
  pr.innerHTML = `<div class="form-preview"><div class="form-preview-title">📋 Anteprima</div><div class="form-preview-row"><span>PM Peptide:</span><span>${pmw.toFixed(2)} Da</span></div>${cycleRow}${preloadedRow}<div class="form-preview-row"><span>Residui:</span><span>${tokens.length}</span></div><div class="form-preview-row"><span>Massa Resina:</span><span>${rm.toFixed(1)} mg</span></div></div>`;
}

function saveSynthesis() {
  const n=document.getElementById('form-name').value.trim();
  const sq=document.getElementById('form-sequence').value.trim();
  const fId=document.getElementById('form-folder-id')?.value||'';
  const dt=document.getElementById('form-date').value;
  const rt=document.getElementById('form-resin').value;
  const rl=parseFloat(document.getElementById('form-loading').value);
  const sc=parseFloat(document.getElementById('form-scale').value);
  const eq=parseFloat(document.getElementById('form-eq').value);
  const a1=document.getElementById('form-activator-1').value;
  const a2=document.getElementById('form-activator-2').value;
  const a3=document.getElementById('form-activator-3').value;
  const preloadedResidue=isWangResinName(rt) ? (document.getElementById('form-preloaded-residue')?.value || '') : '';
  
  if(!n||!sq||!validateSequence()||!rl||!sc||!eq){alert('Verifica i campi obbligatori');return;}
  
  const ns = {
    id:generateId(),
    name:n,
    sequence:sq,
    folderId:fId,
    dateStarted:dt,
    resinType:rt,
    resinLoading:rl,
    scale:sc,
    equivalents:eq,
    preloadedResidue,
    activators:[a1,a2,a3].filter(Boolean),
    activator:a1,
    aaActivators:{},
    customMWs:{},
    status:'in-progress',
    msObserved:'',
    hplcPurity:'',
    hplcRT:'',
    hplcNotes:'',
    notes:'',
    createdAt:new Date().toISOString()
  };
  syntheses.unshift(ns); 
  saveSyntheses(syntheses); 
  navigate('detail', ns.id);
}

function saveEditedSynthesis(synthesisId) {
  const s = syntheses.find(x => x.id === synthesisId);
  if (!s) return;

  const n = document.getElementById('edit-name').value.trim();
  const sq = document.getElementById('edit-sequence').value.trim();
  const fId = document.getElementById('edit-folder-id')?.value || '';
  const dt = document.getElementById('edit-date').value;
  const rt = document.getElementById('edit-resin').value;
  const rl = parseFloat(document.getElementById('edit-loading').value);
  const sc = parseFloat(document.getElementById('edit-scale').value);
  const eq = parseFloat(document.getElementById('edit-eq').value);
  const a1 = document.getElementById('edit-activator-1').value;
  const a2 = document.getElementById('edit-activator-2').value;
  const a3 = document.getElementById('edit-activator-3').value;
  const preloadedResidue = isWangResinName(rt) ? (document.getElementById('edit-preloaded-residue')?.value || '') : '';

  if (!n || !sq || !validateSequenceField('edit') || !rl || !sc || !eq) {
    alert('Verifica i campi obbligatori');
    return;
  }

  const sequenceChanged = s.sequence !== sq;
  const activatorsChanged = JSON.stringify(s.activators || []) !== JSON.stringify([a1, a2, a3].filter(Boolean));

  s.name = n;
  s.sequence = sq;
  s.folderId = fId;
  s.dateStarted = dt;
  s.resinType = rt;
  s.resinLoading = rl;
  s.scale = sc;
  s.equivalents = eq;
  s.preloadedResidue = preloadedResidue;
  s.activators = [a1, a2, a3].filter(Boolean);
  s.activator = a1;

  if (sequenceChanged || activatorsChanged) {
    s.aaActivators = {};
  }

  saveSyntheses(syntheses);
  render();
}

function findInInventory(name) {
  if (!inventoryData || inventoryData.length === 0) return null;
  var n = String(name || '').trim().toLowerCase();
  if (!n) return null;
  var queryAliases = getInventorySearchAliases(name);
  
  // 1. Try Exact/Contains match (priority)
  for (var i = 0; i < inventoryData.length; i++) {
    var row = inventoryData[i];
    if (!row) continue;
    var rowName = String(row[0] || '').toLowerCase();
    var rowCode = String(row[1] || '').toLowerCase();
    var rowCas = String(row[5] || '').toLowerCase();
    if (rowName === n || rowCode === n || rowCas === n || rowName.indexOf(n) !== -1) {
       return { name: row[0], quantity: row[1] || '-', brand: row[2] || '-', code: row[3] || '-', location: row[4] || '-', cas: row[5] || '-', residue: row[6] || '-' };
    }

    var rowAliases = getInventorySearchAliases(row[0]);
    for (var a = 0; a < queryAliases.length; a++) {
      if (rowAliases.indexOf(queryAliases[a]) !== -1) {
        return { name: row[0], quantity: row[1] || '-', brand: row[2] || '-', code: row[3] || '-', location: row[4] || '-', cas: row[5] || '-', residue: row[6] || '-' };
      }
    }
  }

  // 2. Keyword-based fuzzy match for unconventional AA names.
  var queryKeywords = getInventoryKeywordTokens(name);
  var bestMatch = null;
  var bestScore = 0;
  for (var k = 0; k < inventoryData.length; k++) {
    var fuzzyRow = inventoryData[k];
    if (!fuzzyRow) continue;
    var rowKeywords = getInventoryKeywordTokens(fuzzyRow[0]);
    var score = scoreInventoryKeywordMatch(queryKeywords, rowKeywords);
    if (score.score > bestScore) {
      bestScore = score.score;
      bestMatch = { row: fuzzyRow, details: score };
    }
  }

  if (bestMatch && bestMatch.details.sharedCount >= 2 && bestMatch.details.coverage >= 0.5) {
    var row = bestMatch.row;
    return { name: row[0], quantity: row[1] || '-', brand: row[2] || '-', code: row[3] || '-', location: row[4] || '-', cas: row[5] || '-', residue: row[6] || '-' };
  }

  return null;
}

function normalizeInventorySearchTerm(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getInventorySearchAliases(value) {
  var normalized = normalizeInventorySearchTerm(value);
  var aliases = [normalized];

  if (normalized === 'fmoclysallocoh') {
    aliases.push('fmoclysallocoh');
  }
  if (/fmocl?aspoalloh/.test(normalized) || /aspoall/.test(normalized)) {
    aliases.push('fmocaspoalloh', 'fmoclaspoalloh', 'aspoall');
  }

  var unique = [];
  for (var i = 0; i < aliases.length; i++) {
    if (aliases[i] && unique.indexOf(aliases[i]) === -1) unique.push(aliases[i]);
  }
  return unique;
}

function getInventoryKeywordTokens(value) {
  var raw = String(value || '').toLowerCase();
  var expanded = raw
    .replace(/fmoc/g, ' fmoc ')
    .replace(/boc(\d+)/g, ' boc $1 ')
    .replace(/([a-z])(\d+)/g, '$1 $2 ')
    .replace(/(\d+)([a-z])/g, '$1 $2 ')
    .replace(/[^a-z0-9]+/g, ' ');
  var parts = expanded.split(/\s+/).filter(Boolean);
  var stopWords = ['fmoc', 'oh', 'acid', 'amino', 'alloc', 'oall'];
  var aaAliases = {
    phenylalanine: 'phe',
    lysine: 'lys',
    aspartic: 'asp',
    glutamic: 'glu',
    tyrosine: 'tyr',
    tryptophan: 'trp',
    histidine: 'his',
    arginine: 'arg',
    cysteine: 'cys',
    serine: 'ser',
    threonine: 'thr',
    alanine: 'ala',
    glycine: 'gly',
    leucine: 'leu',
    isoleucine: 'ile',
    valine: 'val',
    proline: 'pro',
    methionine: 'met',
    asparagine: 'asn',
    glutamine: 'gln'
  };
  var tokens = [];

  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    if (aaAliases[p]) p = aaAliases[p];
    if (stopWords.indexOf(p) !== -1) continue;
    if (p.length < 2 && !/^\d+$/.test(p)) continue;
    if (tokens.indexOf(p) === -1) tokens.push(p);
  }

  return tokens;
}

function scoreInventoryKeywordMatch(queryKeywords, rowKeywords) {
  if (!queryKeywords.length || !rowKeywords.length) {
    return { score: 0, sharedCount: 0, coverage: 0 };
  }

  var shared = [];
  for (var i = 0; i < queryKeywords.length; i++) {
    if (rowKeywords.indexOf(queryKeywords[i]) !== -1 && shared.indexOf(queryKeywords[i]) === -1) {
      shared.push(queryKeywords[i]);
    }
  }

  var coverage = shared.length / queryKeywords.length;
  var rowCoverage = shared.length / rowKeywords.length;
  var score = shared.length * 10 + coverage * 5 + rowCoverage * 2;
  return { score: score, sharedCount: shared.length, coverage: coverage, rowCoverage: rowCoverage, shared: shared };
}

var INVENTORY_COLOR_WORDS = [
  'BIANCO', 'NERO', 'ROSSO', 'VERDE', 'BLU', 'AZZURRO', 'CELESTE',
  'GIALLO', 'ARANCIONE', 'VIOLA', 'LILLA', 'ROSA', 'GRIGIO',
  'MARRONE', 'TURCHESE', 'BORDEAUX', 'MAGENTA', 'CIANO', 'ORO',
  'ARGENTO'
];
var INVENTORY_COLOR_QUALIFIERS = ['CHIARO', 'SCURO', 'FLUO', 'FLUORESCENTE', 'PASTELLO'];
var INVENTORY_PROTECTING_GROUPS = [
  'ACM', 'ALLOC', 'BOC', 'CBZ', 'DDE', 'IVDDE', 'MBF', 'MMT', 'MTR',
  'MTT', 'OTBU', 'PBF', 'PMC', 'STBU', 'TBU', 'TRT', 'TRITYL', 'Z'
];

function normalizeInventoryLabel(label) {
  return String(label || '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function isInventoryColorLabel(label) {
  var original = String(label || '').replace(/\s+/g, ' ').trim();
  var normalized = normalizeInventoryLabel(original);
  if (!normalized || original !== normalized || !/^[A-Z\s]+$/.test(normalized)) return false;

  var parts = normalized.split(' ');
  var hasColor = false;
  for (var i = 0; i < parts.length; i++) {
    if (INVENTORY_COLOR_WORDS.indexOf(parts[i]) !== -1) {
      hasColor = true;
    } else if (INVENTORY_COLOR_QUALIFIERS.indexOf(parts[i]) === -1) {
      return false;
    }
  }
  return hasColor;
}

function extractInventoryColorFromName(name) {
  var matches = String(name || '').match(/\(([^()]*)\)/g);
  var color = '';
  if (!matches) return color;

  for (var i = 0; i < matches.length; i++) {
    var label = matches[i].slice(1, -1);
    if (isInventoryColorLabel(label)) color = normalizeInventoryLabel(label);
  }
  return color;
}

function isProtectingGroupLabel(label) {
  var compact = String(label || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return INVENTORY_PROTECTING_GROUPS.indexOf(compact) !== -1;
}

function formatInventoryLocation(inv) {
  var location = String(inv.location || '-').trim() || '-';
  var displayLocation = location.replace(/\s*\(([^()]*)\)/g, function(match, label) {
    if (isInventoryColorLabel(label)) return ' (' + normalizeInventoryLabel(label) + ')';
    if (isProtectingGroupLabel(label)) return '';
    return match;
  }).replace(/\s{2,}/g, ' ').trim() || '-';

  var color = extractInventoryColorFromName(inv.name);
  if (color && displayLocation.indexOf('(' + color + ')') === -1) {
    displayLocation += ' (' + color + ')';
  }
  return displayLocation;
}

function renderInventoryCheckRows(aaTokens, activatorNames, synthesis) {
  var combined = aaTokens.concat(activatorNames);
  var uniqueArr = [];
  for (var i = 0; i < combined.length; i++) {
    if (uniqueArr.indexOf(combined[i]) === -1) uniqueArr.push(combined[i]);
  }
  
  return uniqueArr.map(function(item) {
    var aaData = AMINO_ACIDS[item];
    var searchName = aaData ? aaData.fmocName : item;
    var inv = findInInventory(searchName);
    if (inv) {
      var displayLocation = formatInventoryLocation(inv);
      return '<tr><td><strong>' + escapeHtml(searchName) + '</strong></td><td>' + escapeHtml(inv.quantity) + '</td><td>' + escapeHtml(inv.brand) + '</td><td>' + escapeHtml(inv.code) + '</td><td>' + escapeHtml(displayLocation) + '</td><td>' + escapeHtml(inv.cas) + '</td><td>' + escapeHtml(inv.residue) + '</td></tr>';
    } else {
      if (aaData) return '<tr style="opacity:0.6"><td>' + escapeHtml(aaData.name) + '</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td><span style="color:var(--warning)">Non in inventario</span></td></tr>';
      return '<tr><td><strong>' + escapeHtml(item) + '</strong></td><td colspan="5" style="color:var(--danger); text-align:center;">Non trovato in inventario Excel</td><td>-</td></tr>';
    }
  }).join('');
}

async function fetchPubChemMW(cas, token, synthesisId) {
  var s = syntheses.find(function(x) { return x.id === synthesisId; });
  if (!s) return;

  var queries = [];
  if (cas && cas !== '-') queries.push(cas);
  
  // Clean token: remove Fmoc, OH, etc.
  var clean = token.replace(/fmoc|boc|otbu|trt|pbf|\-oh|\(|\)/gi, '').trim();
  if (clean && queries.indexOf(clean) === -1) queries.push(clean);
  if (token && queries.indexOf(token) === -1) queries.push(token);

  for (var i = 0; i < queries.length; i++) {
    try {
      var url = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/' + encodeURIComponent(queries[i]) + '/property/MolecularWeight/JSON';
      var resp = await fetch(url);
      if (resp.ok) {
        var data = await resp.json();
        var mw = data.PropertyTable && data.PropertyTable.Properties && data.PropertyTable.Properties[0] && data.PropertyTable.Properties[0].MolecularWeight;
        if (mw) {
          saveCustomMW(synthesisId, token, mw);
          return;
        }
      }
    } catch (err) { console.error("PubChem fetch error for " + queries[i], err); }
  }
  
  saveCustomMW(synthesisId, token, 'failed');
}

function saveCustomMW(synthesisId, token, mw) {
  var s = syntheses.find(function(x) { return x.id === synthesisId; });
  if (!s) return;
  if (!s.customMWs) s.customMWs = {};
  s.customMWs[token] = (mw === 'failed') ? 'failed' : parseFloat(mw);
  saveSyntheses(syntheses);
  render();
}

function checkMSMatch(s) {
  var mEl = document.getElementById('ms-match');
  if (!mEl) return;
  var obs = parseFloat(s.msObserved);
  if (!obs || isNaN(obs)) { mEl.innerHTML = ''; return; }
  var r = RESINS.find(function(x) { return x.name === s.resinType; });
  var tokens = tokenizeSequence(s.sequence);
  var eMw = calculatePeptideMW(tokens, (r ? r.type : 'amide'), s.customMWs, s.sequence);
  var ad = calculateMSAdducts(eMw);
  var cl = null, mD = Infinity;
  for (var i = 0; i < ad.length; i++) {
    var a = ad[i];
    var d = Math.abs(obs - a.mz);
    if (d < mD) { mD = d; cl = a; }
  }
  if (cl) {
    if (mD < 1.5) mEl.innerHTML = '<span style="color:var(--success)">OK: ' + cl.name + ' = ' + cl.mz.toFixed(2) + ' (diff ' + mD.toFixed(2) + ')</span>';
    else if (mD < 5) mEl.innerHTML = '<span style="color:var(--warning)">PROBABILE: ' + cl.name + ' = ' + cl.mz.toFixed(2) + ' (diff ' + mD.toFixed(2) + ')</span>';
    else mEl.innerHTML = '<span style="color:var(--danger)">NO MATCH (diff ' + mD.toFixed(2) + ')</span>';
  }
}

// ---- Other Helpers ----

function isWangResinName(name) {
  return String(name || '').toLowerCase().indexOf('wang') !== -1;
}

function renderPreloadedResidueOptions(selectedToken) {
  return Object.keys(AMINO_ACIDS).map(function(token) {
    var aa = AMINO_ACIDS[token];
    var selected = token === selectedToken ? 'selected' : '';
    return '<option value="' + token + '" ' + selected + '>' + escapeHtml(aa.fmocName) + ' (' + escapeHtml(aa.code3) + ', ' + token + ')</option>';
  }).join('');
}

function getPreloadedResidueLabel(token) {
  var aa = AMINO_ACIDS[token];
  return aa ? aa.fmocName + ' (' + aa.code3 + ', ' + token + ')' : String(token || '-');
}

function getPreloadedMismatchMessage(preloadedResidue, tokens) {
  if (!preloadedResidue || !tokens || tokens.length === 0) return '';
  var cTerminal = tokens[tokens.length - 1];
  if (String(cTerminal).toUpperCase() === String(preloadedResidue).toUpperCase()) return '';
  return 'Attenzione: il residuo Wang preloaded selezionato non coincide con il C-terminale della sequenza (' + cTerminal + ').';
}

function isPreloadedWangStep(s, token, forwardIdx, totalTokens) {
  return isWangResinName(s && s.resinType) &&
    s.preloadedResidue &&
    forwardIdx === totalTokens - 1 &&
    String(token).toUpperCase() === String(s.preloadedResidue).toUpperCase();
}

function updateWangPreloadedVisibility(prefix) {
  var resinEl = document.getElementById(prefix + '-resin');
  var panel = document.getElementById(prefix + '-wang-preloaded-group');
  var preloadedEl = document.getElementById(prefix + '-preloaded-residue');
  if (!resinEl || !panel) return;
  var show = isWangResinName(resinEl.value);
  panel.style.display = show ? 'block' : 'none';
  if (!show && preloadedEl) preloadedEl.value = '';
}

function renderPreloadedResinWarning(s, tokens) {
  if (!s || !isWangResinName(s.resinType) || !s.preloadedResidue) return '';
  var message = getPreloadedMismatchMessage(s.preloadedResidue, tokens);
  if (!message) return '';
  return '<div class="preloaded-warning animate-in">' + escapeHtml(message) + '</div>';
}

function formatMassDelta(delta) {
  var value = parseFloat(delta) || 0;
  return (value > 0 ? '+' : '') + value.toFixed(3);
}

function renderSequenceWithCyclization(seq) {
  var raw = String(seq || '');
  var html = '';
  var buffer = '';
  var inCycle = false;

  function flushBuffer() {
    if (!buffer) return;
    html += inCycle ? '<span class="sequence-cycle-segment">' + escapeHtml(buffer) + '</span>' : escapeHtml(buffer);
    buffer = '';
  }

  for (var i = 0; i < raw.length; i++) {
    if (raw[i] === '*') {
      flushBuffer();
      html += '<span class="cyclization-marker">*</span>';
      inCycle = !inCycle;
    } else {
      buffer += raw[i];
    }
  }
  flushBuffer();

  return html || escapeHtml(raw);
}

function renderCyclizationSection(info, synthesis) {
  if (!info || info.markerCount === 0) return '';

  var rows = info.cyclizations.map(function(c, idx) {
    var residueLabel = c.startLabel && c.endLabel ? c.startLabel + ' - ' + c.endLabel : '-';
    var rowClass = c.type === 'unknown' ? ' cyclization-row-warning' : '';
    var buildingBlocks = getCyclizationBuildingBlocks(c);
    var buildingBlocksHtml = buildingBlocks.length > 0
      ? '<div class="cyclization-row-detail">Building block: <strong>' + buildingBlocks.map(escapeHtml).join(' + ') + '</strong></div>'
      : '';
    return `
      <div class="cyclization-row${rowClass}">
        <div>
          <div class="cyclization-row-title">${idx + 1}. ${escapeHtml(c.label)}</div>
          <div class="cyclization-row-detail">Porzione: <strong>${escapeHtml(c.segment || '-')}</strong></div>
          ${buildingBlocksHtml}
        </div>
        <div class="cyclization-row-detail">Residui: <strong>${escapeHtml(residueLabel)}</strong></div>
        <div class="cyclization-row-detail mono">Delta: <strong>${formatMassDelta(c.massDelta)} Da</strong></div>
        <div class="cyclization-row-detail">${escapeHtml(c.note || '')}</div>
      </div>`;
  }).join('');

  var warnings = info.warnings.length > 0
    ? '<ul class="cyclization-warnings">' + info.warnings.map(function(w) { return '<li>' + escapeHtml(w) + '</li>'; }).join('') + '</ul>'
    : '';
  var protocol = renderCyclizationProtocol(info, synthesis);

  return `
    <div class="section-card animate-in">
      <div class="section-card-header"><div class="section-card-title">Ciclizzazione Peptidica</div><span class="cyclization-total">Delta totale ${formatMassDelta(info.totalMassDelta)} Da</span></div>
      <div class="section-card-body">
        <div class="cyclization-list">${rows}</div>
        ${protocol}
        ${warnings}
      </div>
    </div>`;
}

function getCyclizationBuildingBlocks(cycle) {
  if (!cycle || cycle.type !== 'amide') return [];
  var blocks = [];
  var startVariant = CYCLIZATION_PROTECTED_AMINO_ACIDS[cycle.startToken];
  var endVariant = CYCLIZATION_PROTECTED_AMINO_ACIDS[cycle.endToken];
  if (startVariant) blocks.push([startVariant.fmocName].concat(startVariant.aliases || []).join(' / '));
  if (endVariant) blocks.push([endVariant.fmocName].concat(endVariant.aliases || []).join(' / '));
  return blocks;
}

function renderCyclizationProtocol(info, synthesis) {
  if (!info || !info.cyclizations || info.cyclizations.length === 0) return '';

  var cards = info.cyclizations.map(function(cycle) {
    var steps = getCyclizationProtocolSteps(cycle);
    var calculations = renderCyclizationProtocolCalculations(cycle, synthesis);
    if (!steps.length) return '';

    return `
      <div class="cyclization-protocol-card">
        <div class="cyclization-protocol-title">Protocollo suggerito: ${escapeHtml(cycle.label)}</div>
        <ol class="cyclization-protocol-list">
          ${steps.map(function(step) {
            return '<li><strong>' + escapeHtml(step.title) + '</strong><span>' + escapeHtml(step.body) + '</span></li>';
          }).join('')}
        </ol>
        ${calculations}
      </div>`;
  }).join('');

  return cards ? '<div class="cyclization-protocol">' + cards + '</div>' : '';
}

function getCyclizationProtocolSteps(cycle) {
  if (!cycle) return [];

  if (cycle.type === 'amide') {
    return [
      {
        title: 'Rimozione Alloc/OAll',
        body: '0,15 eq Pd(PPh3)4, 3 eq DMBA (oppure morfolina se sono presenti Cys), DCM:DMF anidri 3:2. Primo ciclo: 15 min in condizioni anidre.'
      },
      {
        title: 'Lavaggi intermedi e secondo ciclo',
        body: 'Dopo il primo ciclo lavare 3 volte con DMF/DCM, riseccare la resina e ripetere gli stessi step per un secondo ciclo da 15 min.'
      },
      {
        title: 'Lavaggi post-deprotezione',
        body: 'Dopo i 15 min del secondo ciclo lavare 3 volte con DCM e 3 volte con DMF.'
      },
      {
        title: 'Scavenging gruppi allilici',
        body: 'Trattare con ditiocarbammato, 25 mg in 5 mL, per 25 min su shaker.'
      },
      {
        title: 'Controllo Kaiser',
        body: 'Eseguire Kaiser test nel caso di liberazione di gruppi amminici; procedere alla ciclizzazione solo dopo controllo positivo/coerente.'
      },
      {
        title: 'Ciclizzazione lattamica',
        body: 'Attivare il carbossile laterale deprotetto e ciclizzare con PyAOP 3 eq e DIPEA 6 eq in DMF; ripetere se il test di completamento non e soddisfacente.'
      },
      {
        title: 'Controllo finale',
        body: 'Monitorare la ciclizzazione con test colorimetrico quando applicabile e confermare la massa tramite mini-cleavage LC-MS.'
      }
    ];
  }

  if (cycle.type === 'disulfide') {
    return [
      {
        title: 'Liberazione dei tioli',
        body: 'Dopo cleavage/deprotezione, mantenere condizioni adatte alla rimozione dei gruppi protettori della Cys secondo SOP.'
      },
      {
        title: 'Ossidazione disolfuro',
        body: 'Formare il ponte Cys-Cys con condizioni ossidanti controllate e monitorare la conversione tramite LC-MS.'
      }
    ];
  }

  return [
    {
      title: 'Definizione manuale',
      body: 'Tipo di ciclizzazione non riconosciuto automaticamente: definire manualmente deprotezione, attivazione e delta massa prima della sintesi.'
    }
  ];
}

var CYCLIZATION_PROTOCOL_REAGENTS = {
  tetrakis: { name: 'Pd(PPh3)4 / Tetrakis', eq: 0.15, mw: 1155.60 },
  dmba: { name: 'DMBA', eq: 3, mw: 156.14 },
  morpholine: { name: 'Morfolina (Cys presente)', eq: 3, mw: 87.12 },
  pyaop: { name: 'PyAOP', eq: 3, mw: 521.40 },
  dipea: { name: 'DIPEA', eq: 6, mw: 129.24, density: 0.742 },
  dithiocarbamate: { name: 'Ditiocarbammato', fixed: '25 mg in 5 mL, 25 min su shaker' }
};

function getCyclizationSynthesisScale(synthesis) {
  return parseFloat(synthesis && synthesis.scale) || 0;
}

function formatCyclizationScale(scale) {
  var value = parseFloat(scale) || 0;
  return value.toFixed(3).replace(/0+$/, '').replace(/\.$/, '');
}

function formatCyclizationEq(eq) {
  return String(eq).replace('.', ',');
}

function formatCyclizationMass(mg) {
  var value = parseFloat(mg) || 0;
  return (value < 10 ? value.toFixed(2) : value.toFixed(1)) + ' mg';
}

function formatCyclizationVolume(ul) {
  var value = parseFloat(ul) || 0;
  return (value < 10 ? value.toFixed(1) : value.toFixed(0)) + ' µL';
}

function hasCysteineForCyclization(cycle, synthesis) {
  var tokens = [];
  if (synthesis && synthesis.sequence) tokens = tokens.concat(tokenizeSequence(synthesis.sequence));
  if (cycle && cycle.tokens) tokens = tokens.concat(cycle.tokens);
  return tokens.some(function(token) {
    return getCyclizationEndpointType(token) === 'thiol';
  });
}

function buildCyclizationReagentRow(phase, reagent, scale, cycles) {
  if (!reagent || !scale) return null;
  var cycleCount = cycles || 1;
  var perCycleMg = scale * reagent.eq * reagent.mw;
  var totalMg = perCycleMg * cycleCount;
  var amount = cycleCount > 1
    ? formatCyclizationMass(perCycleMg) + ' / ciclo; ' + formatCyclizationMass(totalMg) + ' tot.'
    : formatCyclizationMass(perCycleMg);

  if (reagent.density) {
    var perCycleVolume = perCycleMg / reagent.density;
    var totalVolume = totalMg / reagent.density;
    amount += cycleCount > 1
      ? ' (≈ ' + formatCyclizationVolume(perCycleVolume) + ' / ciclo; ' + formatCyclizationVolume(totalVolume) + ' tot.)'
      : ' (≈ ' + formatCyclizationVolume(perCycleVolume) + ')';
  }

  return {
    phase: phase,
    reagent: reagent.name,
    eq: formatCyclizationEq(reagent.eq),
    mw: reagent.mw.toFixed(2),
    amount: amount
  };
}

function getCyclizationReagentCalculations(cycle, synthesis) {
  if (!cycle || cycle.type !== 'amide') return [];
  var scale = getCyclizationSynthesisScale(synthesis);
  if (!scale) return [];

  var allylScavenger = hasCysteineForCyclization(cycle, synthesis)
    ? CYCLIZATION_PROTOCOL_REAGENTS.morpholine
    : CYCLIZATION_PROTOCOL_REAGENTS.dmba;

  return [
    buildCyclizationReagentRow('Deprotezione Alloc/OAll', CYCLIZATION_PROTOCOL_REAGENTS.tetrakis, scale, 2),
    buildCyclizationReagentRow('Deprotezione Alloc/OAll', allylScavenger, scale, 2),
    {
      phase: 'Scavenging allilici',
      reagent: CYCLIZATION_PROTOCOL_REAGENTS.dithiocarbamate.name,
      eq: '-',
      mw: '-',
      amount: CYCLIZATION_PROTOCOL_REAGENTS.dithiocarbamate.fixed
    },
    buildCyclizationReagentRow('Ciclizzazione', CYCLIZATION_PROTOCOL_REAGENTS.pyaop, scale, 1),
    buildCyclizationReagentRow('Ciclizzazione', CYCLIZATION_PROTOCOL_REAGENTS.dipea, scale, 1)
  ].filter(Boolean);
}

function renderCyclizationProtocolCalculations(cycle, synthesis) {
  var rows = getCyclizationReagentCalculations(cycle, synthesis);
  if (!rows.length) return '';

  var scale = getCyclizationSynthesisScale(synthesis);
  return `
    <div class="cyclization-calculations">
      <div class="cyclization-calculations-title">Calcoli reagenti su scala ${escapeHtml(formatCyclizationScale(scale))} mmol</div>
      <div class="cyclization-calculations-table-wrap">
        <table class="calc-table cyclization-calculations-table">
          <thead>
            <tr><th>Fase</th><th>Reagente</th><th>Eq.</th><th>PM (g/mol)</th><th>Da pesare</th></tr>
          </thead>
          <tbody>
            ${rows.map(function(row) {
              return '<tr><td>' + escapeHtml(row.phase) + '</td><td><strong>' + escapeHtml(row.reagent) + '</strong></td><td class="mw-value">' + escapeHtml(row.eq) + '</td><td class="mw-value">' + escapeHtml(row.mw) + '</td><td class="mass-value">' + escapeHtml(row.amount) + '</td></tr>';
            }).join('')}
          </tbody>
        </table>
      </div>
      <div class="cyclization-calculations-note">Pd(PPh3)4 e DMBA/morfolina sono calcolati per due cicli da 15 min; PyAOP e DIPEA per lo step di ciclizzazione.</div>
    </div>`;
}

function renderCyclizationPrintSection(info, synthesis) {
  if (!info || !info.cyclizations || info.cyclizations.length === 0) return '';

  var blocks = info.cyclizations.map(function(cycle, idx) {
    var steps = getCyclizationProtocolSteps(cycle);
    var calculations = getCyclizationReagentCalculations(cycle, synthesis);
    var buildingBlocks = getCyclizationBuildingBlocks(cycle);
    var residueLabel = cycle.startLabel && cycle.endLabel ? cycle.startLabel + ' - ' + cycle.endLabel : '-';

    return `
      <div class="print-cyclization-block">
        <h5>${idx + 1}. ${escapeHtml(cycle.label)} (${escapeHtml(residueLabel)})</h5>
        <div class="print-cyclization-meta">
          <span>Porzione: ${escapeHtml(cycle.segment || '-')}</span>
          <span>Delta massa: ${formatMassDelta(cycle.massDelta)} Da</span>
          ${buildingBlocks.length ? `<span>Building block: ${buildingBlocks.map(escapeHtml).join(' + ')}</span>` : ''}
        </div>
        ${steps.length ? `<ol>${steps.map(function(step) { return '<li><strong>' + escapeHtml(step.title) + ':</strong> ' + escapeHtml(step.body) + '</li>'; }).join('')}</ol>` : ''}
        ${calculations.length ? `<table class="print-cyclization-calculations"><thead><tr><th>Fase</th><th>Reagente</th><th>Eq</th><th>PM</th><th>Da pesare</th></tr></thead><tbody>${calculations.map(function(row) { return '<tr><td>' + escapeHtml(row.phase) + '</td><td>' + escapeHtml(row.reagent) + '</td><td>' + escapeHtml(row.eq) + '</td><td>' + escapeHtml(row.mw) + '</td><td>' + escapeHtml(row.amount) + '</td></tr>'; }).join('')}</tbody></table>` : ''}
      </div>`;
  }).join('');

  return `
    <div class="print-cyclization">
      <h4>CICLIZZAZIONE PEPTIDICA</h4>
      ${blocks}
    </div>`;
}

function getAvailableActivators(s) { 
  var names = s.activators || [s.activator || ACTIVATORS[0].name]; 
  return names.map(function(n) { return ACTIVATORS.find(function(a) { return a.name === n; }); }).filter(Boolean); 
}

function getStats() { 
  return {
    total: syntheses.length, 
    inProgress: syntheses.filter(function(s) { return s.status === 'in-progress'; }).length, 
    completed: syntheses.filter(function(s) { return ['completed', 'cleaved'].indexOf(s.status) !== -1; }).length, 
    analyzed: syntheses.filter(function(s) { return s.status === 'analyzed'; }).length
  }; 
}

function showSaved(id) { 
  var e = document.getElementById(id); 
  if (e) { e.classList.add('visible'); setTimeout(function() { e.classList.remove('visible'); }, 2000); } 
}

function formatDate(d) { 
  return d ? new Date(d).toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '-'; 
}

function truncateSequence(s, m) { return (s && s.length > m) ? s.slice(0, m) + '...' : s; }
function truncateText(t, m) { return (t && t.length > m) ? t.slice(0, m) + '...' : t; }
function formatFileSize(b) { 
  return b < 1024 ? b + ' B' : b < 1048576 ? (b / 1024).toFixed(1) + ' KB' : (b / 1048576).toFixed(1) + ' MB'; 
}

function escapeHtml(s) { 
  return s ? String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;') : ''; 
}

// ---- Print Logic ----
function printSynthesis(id) {
  const s = syntheses.find(x => x.id === id);
  if (!s) return;

  const resin = RESINS.find(r => r.name === s.resinType) || RESINS[0];
  const forwardTokens = tokenizeSequence(s.sequence);
  const tokens = forwardTokens.slice().reverse(); // Reverse for addition order (C to N)
  const cyclizationInfo = analyzeCyclizations(s.sequence);
  const cyclizationProtectedOverrides = getCyclizationProtectedOverrides(s.sequence);
  const peptideMW = calculatePeptideMW(forwardTokens, resin.type, s.customMWs, s.sequence);
  const resinMass = (s.scale / s.resinLoading) * 1000;

  // Reagent calculations
  const activators = getAvailableActivators(s);
  const diea = activators.find(a => a.name === 'DIEA');
  const otherActs = activators.filter(a => a.name !== 'DIEA');
  
  // Create totals block
  let totalsHtml = '';
  otherActs.forEach(act => {
    const m = s.scale * s.equivalents * act.mw;
    totalsHtml += `<div>${act.name} = ${m.toFixed(1)} mg</div>`;
    if (act.base) {
      const isLiquid = act.base.includes('DIEA') || act.base.includes('DIPEA') || act.base.includes('Collidine');
      const baseMass = s.scale * act.baseEq * act.baseMW;
      if (isLiquid) {
        let density = 0.742; // default DIEA
        if (act.base.includes('Collidine')) density = 0.917;
        totalsHtml += `<div>${act.base} = ${(baseMass / density).toFixed(1)} µL</div>`;
      } else {
        totalsHtml += `<div>${act.base} = ${baseMass.toFixed(1)} mg</div>`;
      }
    }
  });

  if (diea) {
    const m = s.scale * (s.equivalents * 2) * diea.mw;
    const vol = m / 0.742;
    totalsHtml += `<div>DIEA = ${vol.toFixed(1)} µL</div>`;
  }

  const rowsHtml = tokens.map((token, i) => {
    const forwardIdx = forwardTokens.length - 1 - i;
    const protectedOverride = cyclizationProtectedOverrides[forwardIdx];
    const isPreloadedStep = isPreloadedWangStep(s, token, forwardIdx, forwardTokens.length);
    const aaData = protectedOverride ? Object.assign({}, AMINO_ACIDS[token] || {}, protectedOverride) : AMINO_ACIDS[token];
    const name = aaData ? aaData.name : token;
    const mw = aaData ? aaData.fmocMW : (parseFloat(s.customMWs?.[token]) || 0);
    const mass = isPreloadedStep ? 0 : s.scale * s.equivalents * mw;
    const note = isPreloadedStep ? ' (preloaded Wang)' : '';
    return `<tr><td>${i+1}. ${name}${note}</td><td>${mw.toFixed(2)}</td><td>${isPreloadedStep ? '-' : mass.toFixed(1) + ' mg'}</td></tr>`;
  }).join('');

  const cyclizationPrintHtml = cyclizationInfo.cyclizations.length > 0
    ? `<span>CICLIZZAZIONE: ${cyclizationInfo.cyclizations.map(c => escapeHtml(c.label + ' ' + (c.startLabel || '') + '-' + (c.endLabel || ''))).join('; ')}</span><span>DELTA: ${formatMassDelta(cyclizationInfo.totalMassDelta)} Da</span>`
    : '';
  const cyclizationPrintSection = renderCyclizationPrintSection(cyclizationInfo, s);

  const printArea = document.createElement('div');
  printArea.id = 'print-area';
  printArea.innerHTML = `
    <div class="print-header">
      <img src="logo_transparent.png" class="print-logo">
      <table class="print-top-info">
        <tr>
          <td>NOME: ${escapeHtml(s.name)}</td>
          <td>DATA: ${formatDate(new Date())}</td>
          <td>SIGLA: ${s.id.slice(-6).toUpperCase()}</td>
        </tr>
      </table>
    </div>

    <div class="print-sequence-box">
      H-${s.sequence}-${resin.type === 'amide' ? 'NH₂' : 'OH'}
    </div>

    <div class="print-metadata">
      <span>RESINA: ${escapeHtml(s.resinType)} (${resinMass.toFixed(1)} mg)</span>
      <span>LOADING: ${s.resinLoading} mmol/g</span>
      <span>PM ATTESO: ${peptideMW.toFixed(2)} Da</span>
      ${isWangResinName(s.resinType) && s.preloadedResidue ? `<span>WANG PRELOADED: ${escapeHtml(getPreloadedResidueLabel(s.preloadedResidue))}</span>` : ''}
      ${cyclizationPrintHtml}
    </div>

    <div class="print-procedure">
      <h4>PROCEDIMENTO:</h4>
      <ul>
        <li>${s.equivalents} eq. AA</li>
        ${otherActs.map(a => `<li>${s.equivalents} eq. ${a.name}</li>`).join('')}
        ${diea ? `<li>${s.equivalents * 2} eq. DIEA</li>` : ''}
      </ul>
    </div>

    <div class="print-reagents">
      ${totalsHtml}
    </div>

    <table class="print-table">
      <thead>
        <tr>
          <th style="width: 50%">Amminoacido (Ordine di sintesi)</th>
          <th style="width: 20%">P.M.</th>
          <th style="width: 30%">Quantità (mg)</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
      </tbody>
    </table>

    <div class="print-notes">
      <h4>NOTE & OSSERVAZIONI:</h4>
      <div class="print-notes-content">${escapeHtml(s.notes || '') || '<br><br><br><br>'}</div>
    </div>

    ${cyclizationPrintSection}

    <div class="print-footer">
      SPPS Lab Assistant — MoDiLabs Laboratory Sheet
    </div>
  `;

  document.body.appendChild(printArea);
  // Give the browser a moment to render the new content
  setTimeout(() => {
    window.print();
    setTimeout(() => printArea.remove(), 1000);
  }, 100);
}

// ---- Init ----
document.addEventListener('DOMContentLoaded', async function() {
  try {
    await initInventory();
  } catch(e) { console.error('Inventory init failed', e); }
  render();
});
