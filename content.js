// Content script - roda no contexto da página SharePoint

/**
 * Detecta o URL base do site SharePoint (necessário para _api/web/).
 * CRÍTICO: a chamada /_api/web/ deve ser relativa ao site collection,
 * não apenas à origem (domínio).
 */
function getSiteUrl(url) {
  // Melhor fonte: _spPageContextInfo injetado pelo SharePoint na página
  if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) {
    return window._spPageContextInfo.webAbsoluteUrl.replace(/\/$/, '');
  }

  // Tenta extrair o site do parâmetro listurl (common em onedrive.aspx)
  const listurlParam = url.searchParams.get('listurl');
  if (listurlParam) {
    const listPath = decodeURIComponent(listurlParam);
    const m = listPath.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    if (m) return url.origin + m[1];
  }

  // Tenta extrair o site do parâmetro id (caminho relativo ao servidor)
  const idParam = url.searchParams.get('id');
  if (idParam) {
    const folderPath = decodeURIComponent(idParam);
    const m = folderPath.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
    if (m) return url.origin + m[1];
  }

  // Tenta extrair do caminho da URL atual
  const pathM = url.pathname.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
  if (pathM) return url.origin + pathM[1];

  // Fallback: usa a origem como site collection raiz
  return url.origin;
}

/**
 * Detecta o caminho da pasta atual e o site URL.
 * Suporta múltiplos formatos de URL do SharePoint Online.
 */
function getSharePointInfo() {
  const url = new URL(window.location.href);
  const origin = url.origin;
  const siteUrl = getSiteUrl(url);
  let folderPath = null;
  let detectionMethod = '';

  // Formato 1: ?id= (AllItems.aspx, onedrive.aspx moderno)
  const idParam = url.searchParams.get('id');
  if (idParam) {
    folderPath = decodeURIComponent(idParam);
    detectionMethod = '?id';
  }

  // Formato 2: ?RootFolder= (interface clássica)
  if (!folderPath) {
    const rf = url.searchParams.get('RootFolder');
    if (rf) {
      folderPath = decodeURIComponent(rf);
      detectionMethod = '?RootFolder';
    }
  }

  // Formato 3: AllItems.aspx sem parâmetro (raiz da biblioteca)
  // ex: /sites/Site/Documentos/Forms/AllItems.aspx
  if (!folderPath && url.pathname.includes('/Forms/AllItems.aspx')) {
    folderPath = decodeURIComponent(url.pathname.replace('/Forms/AllItems.aspx', ''));
    detectionMethod = 'AllItems.aspx-path';
  }

  // Formato 4: URL direta de pasta (sem /_layouts/ e sem .aspx)
  if (!folderPath) {
    const path = decodeURIComponent(url.pathname);
    if (!path.includes('/_layouts/') &&
        !path.includes('/Forms/') &&
        !path.endsWith('.aspx') &&
        path.match(/\/(?:sites|teams|personal)\//i)) {
      folderPath = path;
      detectionMethod = 'direct-path';
    }
  }

  // Formato 5: _spPageContextInfo.listUrl (fallback para páginas SPFx)
  if (!folderPath && window._spPageContextInfo && window._spPageContextInfo.listUrl) {
    folderPath = window._spPageContextInfo.listUrl;
    detectionMethod = 'spcontext.listUrl';
  }

  // Se o caminho detectado parece ser um arquivo (tem extensão), pega o diretório pai
  if (folderPath) {
    const lastSeg = folderPath.split('/').pop() || '';
    if (lastSeg.includes('.') && !lastSeg.startsWith('.')) {
      folderPath = folderPath.split('/').slice(0, -1).join('/');
      detectionMethod += '+parentDir';
    }
  }

  console.log('[SharePoint Downloader] Contexto detectado:', {
    pageUrl: url.href,
    origin,
    siteUrl,
    folderPath,
    detectionMethod,
    spContextAvailable: !!window._spPageContextInfo
  });

  return { origin, siteUrl, folderPath, detectionMethod };
}

/**
 * Constrói a URL base da REST API para uma pasta.
 * Usa siteUrl (não apenas origin) — crítico para sites /sites/NomeSite.
 */
function buildFolderApiUrl(siteUrl, serverPath) {
  const oDataPath = serverPath.replace(/'/g, "''"); // Escapa aspas simples para OData
  const encodedPath = oDataPath.split('/').map(encodeURIComponent).join('/');
  return `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedPath}')`;
}

/**
 * Busca todas as páginas de uma URL paginada da SharePoint REST API.
 */
async function fetchAllPages(apiUrl) {
  const results = [];
  let nextUrl = apiUrl;

  while (nextUrl) {
    const response = await fetch(nextUrl, {
      headers: { 'Accept': 'application/json;odata=verbose' },
      credentials: 'include'
    });

    if (response.status === 403 || response.status === 401) {
      throw new Error(
        `Acesso negado (${response.status}). Verifique se você tem permissão para acessar esta pasta. ` +
        `URL da API: ${nextUrl}`
      );
    }

    if (!response.ok) {
      const text = await response.text().catch(() => '');
      throw new Error(
        `Erro ${response.status} na API SharePoint.\n` +
        `URL: ${nextUrl}\n` +
        `Resposta: ${text.substring(0, 300)}`
      );
    }

    const data = await response.json();

    if (!data.d || !Array.isArray(data.d.results)) {
      throw new Error(
        `Formato de resposta inesperado da API SharePoint.\n` +
        `URL: ${nextUrl}\n` +
        `Resposta: ${JSON.stringify(data).substring(0, 200)}`
      );
    }

    results.push(...data.d.results);
    nextUrl = data.d.__next || null;
  }

  return results;
}

function sanitizePathSegment(name) {
  return name.replace(/[<>:"|?*\x00-\x1f]/g, '_').replace(/^\.+/, '_').trim();
}

/**
 * Lista todos os arquivos de forma recursiva.
 * IMPORTANTE: usa siteUrl para chamadas de API, origin para URLs de download.
 */
async function listFilesRecursive(siteUrl, origin, folderServerPath, localBasePath, includeSubfolders, progressCallback) {
  const files = [];

  async function processFolder(serverPath, localPath) {
    const folderApiBase = buildFolderApiUrl(siteUrl, serverPath);

    // Lista arquivos nesta pasta
    const filesUrl = `${folderApiBase}/Files?$select=Name,ServerRelativeUrl,Length&$top=5000`;

    let fileItems = [];
    try {
      fileItems = await fetchAllPages(filesUrl);
    } catch (err) {
      console.error(`[SharePoint Downloader] Falha ao listar arquivos em "${serverPath}":`, err);
      if (progressCallback) progressCallback({ type: 'error', message: err.message });
      // Propaga o erro para o primeiro nível (pasta raiz) para que o popup mostre
      if (serverPath === folderServerPath) throw err;
      return; // Subpastas com erro: pula e continua
    }

    for (const file of fileItems) {
      const safeName = sanitizePathSegment(file.Name);
      const filePath = localPath ? `${localPath}/${safeName}` : safeName;
      files.push({
        name: file.Name,
        url: `${origin}${file.ServerRelativeUrl}`,
        localPath: filePath,
        size: parseInt(file.Length, 10) || 0
      });
    }

    if (progressCallback) progressCallback({ type: 'progress', count: files.length });

    if (!includeSubfolders) return;

    // Lista subpastas
    const foldersUrl = `${folderApiBase}/Folders?$select=Name,ServerRelativeUrl&$top=1000`;
    let folderItems = [];
    try {
      folderItems = await fetchAllPages(foldersUrl);
    } catch (err) {
      console.error(`[SharePoint Downloader] Falha ao listar subpastas em "${serverPath}":`, err);
      return;
    }

    for (const folder of folderItems) {
      if (folder.Name === 'Forms' || folder.Name === '_catalogs' || folder.Name === '_cts') continue;
      const safeFolderName = sanitizePathSegment(folder.Name);
      const subLocalPath = localPath ? `${localPath}/${safeFolderName}` : safeFolderName;
      await processFolder(folder.ServerRelativeUrl, subLocalPath);
    }
  }

  await processFolder(folderServerPath, localBasePath);
  return files;
}

// Escuta mensagens do popup/app
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'PING') {
    sendResponse({ ok: true });

  } else if (message.type === 'GET_INFO') {
    try {
      sendResponse(getSharePointInfo());
    } catch (err) {
      sendResponse({ error: err.message });
    }

  } else if (message.type === 'LIST_FILES') {
    listFilesRecursive(
      message.siteUrl,
      message.origin,
      message.folderPath,
      message.localFolder,
      message.includeSubfolders,
      (progress) => {
        chrome.runtime.sendMessage({ type: 'SCAN_PROGRESS', ...progress }).catch(() => {});
      }
    )
      .then(files => sendResponse({ success: true, files }))
      .catch(err => sendResponse({ success: false, error: err.message }));

    return true; // resposta assíncrona
  }
});
