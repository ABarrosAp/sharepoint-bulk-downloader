'use strict';

// ============================================================
// State
// ============================================================

let spTabId     = null;
let spInfo      = null;
let dirHandle   = null;
let fileQueue   = [];
let isRunning   = false;
let isPaused    = false;
let isCancelled = false;

const session = {
  totalFiles: 0, completedFiles: 0, failedFiles: 0, skippedFiles: 0, totalBytes: 0
};

let completedEntries = [];  // { name, localPath, size, status, error, time }
let failedEntries    = [];  // same shape — used for retry
let logEntries       = [];  // { time, level, message, detail }
let statsInterval    = null;

// ============================================================
// Speed Tracker  (5-second rolling window)
// ============================================================

const speedTracker = (() => {
  const WINDOW = 5000;
  let samples   = [];   // [{ time, cumBytes }]
  let cumBytes  = 0;
  let startTime = null;

  return {
    reset() { samples = []; cumBytes = 0; startTime = Date.now(); },
    record(bytes) {
      cumBytes += bytes;
      const now = Date.now();
      samples.push({ time: now, bytes: cumBytes });
      const cutoff = now - WINDOW;
      while (samples.length > 1 && samples[0].time < cutoff) samples.shift();
    },
    speed() {
      if (samples.length < 2) return 0;
      const f = samples[0], l = samples[samples.length - 1];
      const dt = (l.time - f.time) / 1000;
      return dt > 0.2 ? (l.bytes - f.bytes) / dt : 0;
    },
    avgSpeed() {
      if (!startTime || cumBytes === 0) return 0;
      const dt = (Date.now() - startTime) / 1000;
      return dt > 0 ? cumBytes / dt : 0;
    },
    total()   { return cumBytes; },
    elapsed() { return startTime ? (Date.now() - startTime) / 1000 : 0; }
  };
})();

// ============================================================
// Log module
// ============================================================

function addLog(level, message, detail = '') {
  const entry = { time: Date.now(), level, message, detail };
  logEntries.push(entry);
  renderLogEntry(entry);
  setText('badge-log', logEntries.length);
  setText('log-count', `${logEntries.length} entrada(s)`);
}

function renderLogEntry(entry) {
  const list = document.getElementById('log-list');
  if (!list) return;

  const d = document.createElement('div');
  d.className = `log-entry log-${entry.level}`;
  const t = new Date(entry.time).toLocaleTimeString('pt-BR');
  d.innerHTML =
    `<span class="log-time">${t}</span>` +
    `<span class="log-lvl">${entry.level.toUpperCase()}</span>` +
    `<span class="log-msg">${esc(entry.message)}</span>` +
    (entry.detail ? `<span class="log-detail">${esc(entry.detail)}</span>` : '');
  list.appendChild(d);

  // Auto-scroll if near bottom
  if (list.scrollHeight - list.scrollTop - list.clientHeight < 80) {
    list.scrollTop = list.scrollHeight;
  }

  // Keep DOM lean (max 600 rows visible)
  while (list.childElementCount > 600) list.removeChild(list.firstChild);
}

// ============================================================
// Completed / Failed entries
// ============================================================

function addCompleted(file, status, error = '') {
  const entry = { name: file.name, localPath: file.localPath, size: file.size || 0, status, error, time: Date.now() };
  completedEntries.push(entry);
  if (status === 'failed') failedEntries.push({ ...file });
  renderCompletedEntry(entry);

  setText('badge-completed', completedEntries.length);
  updateCompletedSummary();

  const retryBtn = document.getElementById('btn-retry-failed');
  if (retryBtn && failedEntries.length > 0) {
    retryBtn.classList.remove('hidden');
    retryBtn.textContent = `Tentar novamente (${failedEntries.length} falha${failedEntries.length > 1 ? 's' : ''})`;
  }
}

function renderCompletedEntry(entry) {
  const list = document.getElementById('completed-list');
  if (!list) return;

  const icon    = { ok: '✓', failed: '✗', skipped: '–' }[entry.status];
  const cls     = { ok: 'ci-ok', failed: 'ci-fail', skipped: 'ci-skip' }[entry.status];
  const t       = new Date(entry.time).toLocaleTimeString('pt-BR');

  const d = document.createElement('div');
  d.className = `completed-item ${cls}`;
  d.title = entry.status === 'failed' ? `ERRO: ${entry.error}` : entry.localPath;
  d.innerHTML =
    `<span class="ci-icon">${icon}</span>` +
    `<span class="ci-path">${esc(entry.localPath)}</span>` +
    `<span class="ci-meta">${entry.size > 0 ? formatSize(entry.size) : ''} <span class="ci-time">${t}</span></span>`;
  list.appendChild(d);

  if (isRunning) list.scrollTop = list.scrollHeight;
}

function updateCompletedSummary() {
  const ok      = completedEntries.filter(e => e.status === 'ok').length;
  const failed  = completedEntries.filter(e => e.status === 'failed').length;
  const skipped = completedEntries.filter(e => e.status === 'skipped').length;
  setText('completed-summary',
    `${ok} ok${failed > 0 ? ` · ${failed} falha(s)` : ''}${skipped > 0 ? ` · ${skipped} ignorado(s)` : ''}`
  );
}

// ============================================================
// Boot
// ============================================================

document.addEventListener('DOMContentLoaded', async () => {
  setupButtons();
  await loadSharePointTabs();
  await tryRestoreDirectory();
});

function setupButtons() {
  on('btn-refresh-tabs',  loadSharePointTabs);
  on('btn-pick-dir',      pickDirectory);
  on('btn-scan',          scanFiles);
  on('btn-start',         startDownloads);
  on('btn-pause',         pauseDownloads);
  on('btn-resume',        resumeDownloads);
  on('btn-cancel',        cancelDownloads);
  on('btn-change-tab',    showTabSelector);
  on('btn-manual-apply',  applyManualPath);
  on('btn-export-log',    exportLog);
  on('btn-clear-log',     clearLog);
  on('btn-retry-failed',  retryFailed);

  document.querySelectorAll('.tab-nav-btn').forEach(btn =>
    btn.addEventListener('click', () => switchTab(btn.dataset.tab))
  );
}

// ============================================================
// Tab navigation
// ============================================================

function switchTab(name) {
  document.querySelectorAll('.tab-nav-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.toggle('hidden', p.id !== `tab-${name}`));
}

// ============================================================
// SharePoint tab connection
// ============================================================

async function loadSharePointTabs() {
  const tabs = await chrome.tabs.query({ url: 'https://*.sharepoint.com/*' });
  const listEl = document.getElementById('sp-tab-list');
  listEl.innerHTML = '';

  if (tabs.length === 0) { show('sp-tab-none'); return; }
  hide('sp-tab-none');

  const { lastSpTabId } = await chrome.storage.session.get('lastSpTabId');
  for (const tab of tabs) {
    const btn = document.createElement('button');
    btn.className = 'tab-btn' + (tab.id === lastSpTabId ? ' preferred' : '');
    btn.textContent = tab.title || tab.url;
    btn.title = tab.url;
    btn.addEventListener('click', () => connectToTab(tab));
    listEl.appendChild(btn);
  }

  const preferred = tabs.find(t => t.id === lastSpTabId) || (tabs.length === 1 ? tabs[0] : null);
  if (preferred) connectToTab(preferred);
}

async function connectToTab(tab) {
  spTabId = tab.id;
  hide('sp-tab-area');
  show('sp-connected');

  try {
    await ensureContentScript(tab.id);
    const info = await sendToContent(tab.id, { type: 'GET_INFO' });

    if (!info?.folderPath) {
      document.getElementById('folder-path').textContent = 'Pasta não detectada';
      document.getElementById('folder-path').style.color = '#a4262c';
      show('manual-path-panel');
      setStatus('Pasta não detectada. Informe o caminho manualmente.', 'warn');
    } else {
      spInfo = info;
      document.getElementById('folder-path').textContent = info.folderPath;
      document.getElementById('folder-path').style.color = '';
      document.getElementById('site-url').textContent = info.siteUrl || info.origin;
      hide('manual-path-panel');
      const suggested = info.folderPath.split('/').filter(Boolean).pop() || 'SharePoint-Download';
      if (!document.getElementById('local-folder').value) {
        document.getElementById('local-folder').value = suggested;
      }
      setStatus('Conectado. Escolha a pasta de destino e liste os arquivos.');
    }
  } catch (err) {
    document.getElementById('folder-path').textContent = 'Erro ao comunicar com a aba';
    document.getElementById('folder-path').style.color = '#a4262c';
    show('manual-path-panel');
    setStatus(`Erro: ${err.message}`, 'error');
  }

  checkReadyToScan();
}

function showTabSelector() { hide('sp-connected'); show('sp-tab-area'); loadSharePointTabs(); }

async function applyManualPath() {
  let input = document.getElementById('manual-folder-path').value.trim();
  if (!input) return;

  let folderPath = input;
  let siteUrl = spInfo?.siteUrl || spInfo?.origin || '';

  try {
    if (input.startsWith('http')) {
      const u = new URL(input);
      folderPath = u.searchParams.get('id') || u.searchParams.get('RootFolder') || decodeURIComponent(u.pathname);
      if (!siteUrl) {
        const m = folderPath.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
        siteUrl = m ? u.origin + m[1] : u.origin;
      }
    }
  } catch (_) {}

  if (!siteUrl && spInfo?.origin) {
    const m = folderPath.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    siteUrl = m ? spInfo.origin + m[1] : spInfo.origin;
  }

  spInfo = { ...(spInfo || {}), folderPath, siteUrl };
  document.getElementById('folder-path').textContent = folderPath;
  document.getElementById('folder-path').style.color = '';
  document.getElementById('site-url').textContent = siteUrl;
  hide('manual-path-panel');
  setStatus('Caminho manual aplicado.', 'ok');
  checkReadyToScan();
}

// ============================================================
// Directory selection  (File System Access API)
// ============================================================

async function pickDirectory() {
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'downloads' });
    dirHandle = handle;
    await saveHandleToIDB(handle);
    updateDirUI(handle.name);
    checkReadyToScan();
    setStatus('Pasta de destino: ' + handle.name);
  } catch (err) {
    if (err.name !== 'AbortError') setStatus('Erro ao selecionar pasta: ' + err.message, 'error');
  }
}

async function tryRestoreDirectory() {
  try {
    const handle = await loadHandleFromIDB();
    if (!handle) return;
    if ((await handle.queryPermission({ mode: 'readwrite' })) === 'granted') {
      dirHandle = handle;
      updateDirUI(handle.name);
      checkReadyToScan();
    }
  } catch (_) {}
}

function updateDirUI(name) {
  document.getElementById('dir-name').textContent = name;
  show('dir-info'); hide('no-dir-msg');
  updateSubfolderHint();
}

function updateSubfolderHint() {
  const sub = document.getElementById('local-folder').value.trim() || '(raiz)';
  document.getElementById('dir-subfolder').textContent = sub;
}

document.addEventListener('input', e => { if (e.target.id === 'local-folder') updateSubfolderHint(); });

// ============================================================
// File scanning
// ============================================================

async function scanFiles() {
  if (!spInfo?.folderPath || !spTabId) {
    setStatus('Conecte a uma aba do SharePoint primeiro.', 'error');
    return;
  }

  const localFolder = document.getElementById('local-folder').value.trim() || 'SharePoint-Download';
  const includeSubs = document.getElementById('include-subfolders').checked;

  setScanLoading(true);
  hide('panel-files');
  setText('badge-queue', '0');
  setStatus('Listando arquivos...');

  const onScanProgress = msg => {
    if (msg.type === 'SCAN_PROGRESS' && msg.count !== undefined) {
      setStatus(`Escaneando... ${msg.count} arquivo(s) encontrados`);
      setText('badge-queue', msg.count);
    }
  };
  chrome.runtime.onMessage.addListener(onScanProgress);

  try {
    const result = await sendToContent(spTabId, {
      type: 'LIST_FILES',
      siteUrl: spInfo.siteUrl, origin: spInfo.origin,
      folderPath: spInfo.folderPath, localFolder, includeSubfolders: includeSubs
    });

    if (!result?.success) throw new Error(result?.error || 'Falha desconhecida');

    fileQueue = result.files;
    renderFileList(fileQueue);
    setStatus(`${fileQueue.length} arquivo(s) encontrado(s). Pronto para baixar.`, 'ok');
    enable('btn-start');

    addLog('info', `Listagem concluída: ${fileQueue.length} arquivo(s)`,
      `Pasta: ${spInfo.folderPath} · Subpastas: ${includeSubs ? 'sim' : 'não'}`);
  } catch (err) {
    setStatus('Erro ao listar: ' + err.message, 'error');
    addLog('error', 'Falha ao listar arquivos', err.message);
  } finally {
    chrome.runtime.onMessage.removeListener(onScanProgress);
    setScanLoading(false);
  }
}

function renderFileList(files) {
  setText('file-count', files.length);
  const totalBytes = files.reduce((s, f) => s + (f.size || 0), 0);
  document.getElementById('total-size').textContent = totalBytes > 0 ? `(${formatSize(totalBytes)} total)` : '';

  const listEl = document.getElementById('file-list');
  listEl.innerHTML = '';
  const MAX = 500;
  files.slice(0, MAX).forEach(f => {
    const d = document.createElement('div');
    d.className = 'file-item';
    d.textContent = f.localPath;
    d.title = f.url;
    listEl.appendChild(d);
  });
  if (files.length > MAX) {
    const d = document.createElement('div');
    d.className = 'file-item more';
    d.textContent = `... e mais ${files.length - MAX} arquivo(s) não exibidos`;
    listEl.appendChild(d);
  }

  hide('panel-empty');
  show('panel-files');
  setText('badge-queue', files.length);
  switchTab('queue');
}

// ============================================================
// Downloads
// ============================================================

async function startDownloads() {
  if (!dirHandle) { setStatus('Selecione uma pasta de destino primeiro.', 'error'); return; }
  if (!fileQueue.length) return;

  if ((await dirHandle.requestPermission({ mode: 'readwrite' })) !== 'granted') {
    setStatus('Permissão de escrita negada. Selecione a pasta novamente.', 'error');
    dirHandle = null; hide('dir-info'); show('no-dir-msg'); return;
  }

  const skipExisting = document.getElementById('skip-existing').checked;
  const queue = [...fileQueue];

  // Reset session state
  session.totalFiles    = queue.length;
  session.completedFiles = 0;
  session.failedFiles   = 0;
  session.skippedFiles  = 0;
  session.totalBytes    = queue.reduce((s, f) => s + (f.size || 0), 0);

  completedEntries = []; failedEntries = [];
  document.getElementById('completed-list').innerHTML = '';
  document.getElementById('log-list').innerHTML = '';
  logEntries = [];
  setText('badge-completed', '0');
  setText('badge-log', '0');
  setText('log-count', '0 entradas');
  hide('btn-retry-failed');
  updateCompletedSummary();

  speedTracker.reset();
  isRunning = true; isPaused = false; isCancelled = false;

  disable('btn-scan'); disable('btn-start');
  show('panel-progress'); show('btn-pause'); show('btn-cancel');
  hide('panel-empty');

  startStatsInterval();

  addLog('info', 'Download iniciado',
    `${queue.length} arquivo(s) · ${formatSize(session.totalBytes)} · Destino: ${dirHandle.name}`);
  setStatus(`Baixando para ${dirHandle.name}...`);

  for (let i = 0; i < queue.length; i++) {
    if (isCancelled) break;
    while (isPaused) await delay(300);
    if (isCancelled) break;

    const file = queue[i];
    updateCurrentFile(file.name, i + 1, queue.length);

    try {
      if (skipExisting && await fileExistsInDir(dirHandle, file.localPath)) {
        session.skippedFiles++;
        session.completedFiles++;
        addCompleted(file, 'skipped');
        addLog('info', `Ignorado (já existe): ${file.name}`);
      } else {
        await downloadAndWrite(dirHandle, file.url, file.localPath, bytes => speedTracker.record(bytes));
        session.completedFiles++;
        addCompleted(file, 'ok');
        addLog('success', `OK: ${file.name}`, formatSize(file.size || 0));
      }
    } catch (err) {
      session.failedFiles++;
      session.completedFiles++;
      addCompleted(file, 'failed', err.message);
      addLog('error', `FALHA: ${file.name}`, err.message);
    }

    updateOverallProgress();
  }

  stopStatsInterval();
  isRunning = false;
  onComplete();
}

async function downloadAndWrite(dirHandle, url, localPath, onChunk) {
  const response = await fetch(url, { credentials: 'include' });
  if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);

  const contentLength = parseInt(response.headers.get('content-length') || '0', 10);

  // Navigate/create directory structure
  const parts = localPath.split('/').filter(Boolean);
  let cur = dirHandle;
  for (let i = 0; i < parts.length - 1; i++) {
    cur = await cur.getDirectoryHandle(parts[i], { create: true });
  }
  const fileHandle = await cur.getFileHandle(parts[parts.length - 1], { create: true });
  const writable   = await fileHandle.createWritable();

  if (response.body) {
    let fileBytes = 0;
    const progressStream = new TransformStream({
      transform(chunk, controller) {
        fileBytes += chunk.byteLength;
        onChunk(chunk.byteLength);
        updateFileProgress(fileBytes, contentLength);
        controller.enqueue(chunk);
      }
    });
    await response.body.pipeThrough(progressStream).pipeTo(writable);
  } else {
    const blob = await response.blob();
    onChunk(blob.size);
    updateFileProgress(blob.size, blob.size);
    await writable.write(blob);
    await writable.close();
  }
}

async function fileExistsInDir(dirHandle, localPath) {
  const parts = localPath.split('/').filter(Boolean);
  let cur = dirHandle;
  try {
    for (let i = 0; i < parts.length - 1; i++) cur = await cur.getDirectoryHandle(parts[i]);
    await cur.getFileHandle(parts[parts.length - 1]);
    return true;
  } catch (_) { return false; }
}

// ============================================================
// Download controls
// ============================================================

function pauseDownloads() {
  isPaused = true;
  hide('btn-pause'); show('btn-resume');
  setText('progress-title', 'Pausado');
  setStatus('Pausado — o arquivo atual será concluído antes de parar.');
  addLog('warn', 'Download pausado');
}

function resumeDownloads() {
  isPaused = false;
  hide('btn-resume'); show('btn-pause');
  setText('progress-title', 'Baixando...');
  addLog('info', 'Download retomado');
}

function cancelDownloads() {
  if (!confirm('Cancelar todos os downloads restantes?')) return;
  isCancelled = true; isPaused = false; isRunning = false;
  stopStatsInterval();
  hide('btn-pause'); hide('btn-resume'); hide('btn-cancel');
  enable('btn-scan'); enable('btn-start');
  setText('progress-title', 'Cancelado');
  setText('current-file-label', '');
  hide('file-progress-wrap');
  addLog('warn', 'Download cancelado',
    `${session.completedFiles} processados · ${session.failedFiles} falhas`);
  setStatus(`Cancelado. ${session.completedFiles} arquivo(s) processado(s).`);
}

function onComplete() {
  stopStatsInterval();
  hide('btn-pause'); hide('btn-resume'); hide('btn-cancel');
  enable('btn-scan');

  document.getElementById('progress-fill').style.width = '100%';
  document.getElementById('progress-fill').classList.add('done');
  setText('current-file-label', 'Concluído!');
  hide('file-progress-wrap');

  const { totalFiles, failedFiles, skippedFiles } = session;
  const ok = totalFiles - failedFiles - skippedFiles;

  setText('progress-title', failedFiles > 0 ? `Concluído com ${failedFiles} falha(s)` : 'Concluído!');

  addLog(failedFiles > 0 ? 'warn' : 'success', 'Download concluído',
    `${ok} ok · ${failedFiles} falhas · ${skippedFiles} ignorados · ` +
    `Tempo: ${formatDuration(speedTracker.elapsed())} · Vel. média: ${formatSpeed(speedTracker.avgSpeed())}`);

  updateStatsUI(true);
  setStatus(failedFiles > 0
    ? `Concluído com ${failedFiles} falha(s). ${ok} baixados.`
    : `Todos os ${totalFiles} arquivo(s) baixados com sucesso!`, failedFiles > 0 ? 'warn' : 'ok');

  switchTab('completed');
}

async function retryFailed() {
  if (!failedEntries.length) return;
  fileQueue = [...failedEntries];
  addLog('info', `Retentando ${fileQueue.length} arquivo(s) com falha`);
  await startDownloads();
}

// ============================================================
// Stats interval — updates UI every 500ms during download
// ============================================================

function startStatsInterval() {
  stopStatsInterval();
  statsInterval = setInterval(() => updateStatsUI(false), 500);
}

function stopStatsInterval() {
  if (statsInterval) { clearInterval(statsInterval); statsInterval = null; }
}

function updateStatsUI(final = false) {
  const speed   = speedTracker.speed();
  const avg     = speedTracker.avgSpeed();
  const elapsed = speedTracker.elapsed();
  const downloaded = speedTracker.total();
  const remaining  = Math.max(0, session.totalBytes - downloaded);
  const eta = (speed > 0 && remaining > 0) ? remaining / speed : null;

  const done  = session.completedFiles;
  const total = session.totalFiles;
  const pct   = total > 0 ? Math.round((done / total) * 100) : 0;

  if (!final) {
    document.getElementById('progress-fill').style.width = `${pct}%`;
    setText('pct-label', `${pct}%`);
  }

  // Stats grid
  setText('stat-speed',      final ? '—' : formatSpeed(speed));
  setText('stat-eta',        final ? 'Concluído' : (eta !== null ? '~' + formatDuration(eta) : '—'));
  setText('stat-elapsed',    formatDuration(elapsed));
  setText('stat-downloaded', `${formatSize(downloaded)} / ${formatSize(session.totalBytes)}`);
  setText('stat-files',      `${done} / ${total}`);
  setText('stat-avg-speed',  formatSpeed(avg));

  // Summary chips
  const ok     = done - session.failedFiles;
  const failed = session.failedFiles;
  setText('stats-completed', `${ok} ok${session.skippedFiles > 0 ? ` (${session.skippedFiles} ignorados)` : ''}`);
  setText('stats-remaining', `${total - done} restantes`);
  const failEl = document.getElementById('stats-failed');
  if (failed > 0) { setText('stats-failed', `${failed} falha(s)`); failEl.classList.remove('hidden'); }
}

function updateCurrentFile(name, idx, total) {
  setText('progress-title', `Baixando ${idx} de ${total}...`);
  setText('current-file-label', name);
  document.getElementById('file-progress-fill').style.width = '0%';
  setText('file-pct-label', '');
  show('file-progress-wrap');
}

function updateFileProgress(bytes, total) {
  if (!total) return;
  const pct = Math.min(100, Math.round((bytes / total) * 100));
  document.getElementById('file-progress-fill').style.width = `${pct}%`;
  setText('file-pct-label', `${formatSize(bytes)} / ${formatSize(total)} · ${pct}%`);
}

function updateOverallProgress() { updateStatsUI(false); }

// ============================================================
// Log export / clear
// ============================================================

function exportLog() {
  if (!logEntries.length) { alert('Log vazio.'); return; }

  const header = [
    'SharePoint Bulk Downloader — Log de Download',
    `Gerado em: ${new Date().toLocaleString('pt-BR')}`,
    `Pasta SharePoint: ${spInfo?.folderPath || '—'}`,
    `Pasta destino: ${dirHandle?.name || '—'}`,
    `Total de arquivos: ${session.totalFiles} · Falhas: ${session.failedFiles}`,
    '─'.repeat(80)
  ].join('\n');

  const lines = logEntries.map(e => {
    const t  = new Date(e.time).toLocaleString('pt-BR');
    const lv = e.level.toUpperCase().padEnd(7);
    return `[${t}] [${lv}] ${e.message}${e.detail ? '\n             → ' + e.detail : ''}`;
  });

  const blob = new Blob([header + '\n\n' + lines.join('\n')], { type: 'text/plain;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), {
    href: url,
    download: `sp-download-log-${new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-')}.txt`
  });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function clearLog() {
  if (!confirm('Limpar o log?')) return;
  logEntries = [];
  document.getElementById('log-list').innerHTML = '';
  setText('log-count', '0 entradas');
  setText('badge-log', '0');
}

// ============================================================
// Helpers
// ============================================================

function checkReadyToScan() {
  const ready = !!spTabId && !!spInfo?.folderPath && !!dirHandle;
  document.getElementById('btn-scan').disabled = !ready;
  if (!ready) disable('btn-start');
}

function setScanLoading(loading) {
  document.getElementById('btn-scan').disabled = loading;
  document.getElementById('scan-spinner').classList.toggle('hidden', !loading);
  setText('scan-label', loading ? 'Listando...' : 'Listar Arquivos');
}

function setStatus(msg, type = '') {
  const el = document.getElementById('status-msg');
  el.textContent = msg; el.className = type;
}

function setText(id, val) { const e = document.getElementById(id); if (e) e.textContent = val; }

function formatSize(bytes) {
  if (!bytes) return '0 B';
  const u = ['B','KB','MB','GB','TB'];
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), 4);
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${u[i]}`;
}

function formatSpeed(bps) {
  if (!bps || bps < 100) return '0 KB/s';
  if (bps < 1024 * 1024) return `${(bps / 1024).toFixed(1)} KB/s`;
  return `${(bps / (1024 * 1024)).toFixed(2)} MB/s`;
}

function formatDuration(s) {
  if (!s || s < 1) return '< 1s';
  if (s < 60)   return `${Math.floor(s)}s`;
  if (s < 3600) return `${Math.floor(s / 60)}m ${Math.floor(s % 60)}s`;
  return `${Math.floor(s / 3600)}h ${Math.floor((s % 3600) / 60)}m`;
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function show(id)   { document.getElementById(id)?.classList.remove('hidden'); }
function hide(id)   { document.getElementById(id)?.classList.add('hidden'); }
function enable(id) { const e = document.getElementById(id); if (e) e.disabled = false; }
function disable(id){ const e = document.getElementById(id); if (e) e.disabled = true; }
function on(id, fn) { document.getElementById(id)?.addEventListener('click', fn); }
function delay(ms)  { return new Promise(r => setTimeout(r, ms)); }

function sendToContent(tabId, message) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, message, r => {
      if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
      else resolve(r);
    });
  });
}

async function ensureContentScript(tabId) {
  try { await sendToContent(tabId, { type: 'PING' }); }
  catch (_) {
    await chrome.scripting.executeScript({ target: { tabId }, files: ['content.js'] });
    await delay(300);
  }
}

// ============================================================
// IndexedDB — persist FileSystemDirectoryHandle
// ============================================================

const IDB_NAME = 'sp-downloader-v2', IDB_STORE = 'handles';

function openIDB() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(IDB_NAME, 1);
    r.onupgradeneeded = e => e.target.result.createObjectStore(IDB_STORE);
    r.onsuccess = e => res(e.target.result);
    r.onerror   = e => rej(e.target.error);
  });
}

async function saveHandleToIDB(handle) {
  const db = await openIDB();
  await new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).put(handle, 'dir');
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
  db.close();
}

async function loadHandleFromIDB() {
  const db = await openIDB();
  const h = await new Promise((res, rej) => {
    const r = db.transaction(IDB_STORE).objectStore(IDB_STORE).get('dir');
    r.onsuccess = () => res(r.result || null); r.onerror = () => rej(r.error);
  });
  db.close();
  return h;
}
