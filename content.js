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
 * Listagem via RenderListDataAsStream.
 *
 * Por que não GetFolderByServerRelativeUrl(...)/Files: essa consulta é bloqueada
 * pelo limite do modo de exibição de lista (5.000 itens, SPQueryThrottledException)
 * quando a pasta passa desse tamanho. RenderListDataAsStream é o mecanismo usado pela
 * própria interface do SharePoint: ordena pelo ID (indexado) e pagina via NextHref,
 * o que mantém cada consulta abaixo do limite.
 */

// Tamanho de página. Precisa ficar abaixo do limite de 5.000 itens.
const LIST_PAGE_SIZE = 4000;

// Pastas de sistema da biblioteca que nunca devem ser baixadas.
const SYSTEM_FOLDERS = new Set(['Forms', '_catalogs', '_cts']);

// Margem de segurança antes de renovar o form digest (ms).
const DIGEST_SAFETY_MARGIN_MS = 60 * 1000;

const LIST_VIEW_XML =
  '<View>' +
    '<Query><OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy></Query>' +
    '<ViewFields>' +
      '<FieldRef Name="FileLeafRef"/>' +
      '<FieldRef Name="FileRef"/>' +
      '<FieldRef Name="FSObjType"/>' +
      '<FieldRef Name="File_x0020_Size"/>' +
    '</ViewFields>' +
    `<RowLimit Paged="TRUE">${LIST_PAGE_SIZE}</RowLimit>` +
  '</View>';

/**
 * fetch() com a sessão do navegador e tratamento de erro padronizado.
 * Retorna o JSON da resposta.
 */
async function spFetchJson(url, options = {}) {
  const response = await fetch(url, {
    credentials: 'include',
    ...options,
    headers: { 'Accept': 'application/json;odata=nometadata', ...(options.headers || {}) }
  });

  if (response.status === 403 || response.status === 401) {
    throw new Error(
      `Acesso negado (${response.status}). Verifique se você tem permissão para acessar esta pasta. ` +
      `URL da API: ${url}`
    );
  }

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(
      `Erro ${response.status} na API SharePoint.\n` +
      `URL: ${url}\n` +
      `Resposta: ${text.substring(0, 300)}`
    );
  }

  return response.json();
}

/**
 * Cria um provedor de form digest (exigido em POST na REST API) com cache,
 * renovado automaticamente antes de expirar. Varreduras longas passam de 30 min.
 */
function createDigestProvider(siteUrl) {
  let value = null;
  let expiresAt = 0;

  return async function getDigest() {
    if (value && Date.now() < expiresAt) return value;
    const data = await spFetchJson(`${siteUrl}/_api/contextinfo`, { method: 'POST' });
    value = data.FormDigestValue;
    expiresAt = Date.now() + data.FormDigestTimeoutSeconds * 1000 - DIGEST_SAFETY_MARGIN_MS;
    return value;
  };
}

const normalizePath = p => decodeURIComponent(p).replace(/\/+$/, '').toLowerCase();

/**
 * Descobre a biblioteca que contém a pasta: a lista cuja RootFolder é o
 * prefixo mais longo do caminho. Consulta a coleção de listas, não os itens,
 * então não sofre com o limite de 5.000.
 */
async function findLibraryForFolder(siteUrl, folderServerPath) {
  const data = await spFetchJson(
    `${siteUrl}/_api/web/lists?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder&$top=5000`
  );
  const target = normalizePath(folderServerPath);

  const library = data.value
    .map(list => ({ id: list.Id, title: list.Title, root: normalizePath(list.RootFolder.ServerRelativeUrl) }))
    .filter(list => target === list.root || target.startsWith(list.root + '/'))
    .sort((a, b) => b.root.length - a.root.length)[0];

  if (!library) {
    throw new Error(
      `Nenhuma biblioteca do site contém a pasta informada.\n` +
      `Site: ${siteUrl}\n` +
      `Pasta: ${folderServerPath}`
    );
  }
  return library;
}

/**
 * Lista os filhos diretos (arquivos e subpastas) de uma pasta, página por página.
 */
async function listFolderChildren(siteUrl, listId, folderServerPath, getDigest) {
  const endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/RenderListDataAsStream`;
  const body = JSON.stringify({
    parameters: {
      RenderOptions: 2, // ListData: só as linhas, sem esquema/HTML
      FolderServerRelativeUrl: folderServerPath,
      ViewXml: LIST_VIEW_XML
    }
  });

  const rows = [];
  let query = '';
  do {
    const data = await spFetchJson(endpoint + query, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': await getDigest()
      },
      body
    });
    rows.push(...(data.Row || []));
    query = data.NextHref || ''; // "?Paged=TRUE&p_ID=...", vazio na última página
  } while (query);

  return rows;
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
  const getDigest = createDigestProvider(siteUrl);
  const library = await findLibraryForFolder(siteUrl, folderServerPath);

  async function processFolder(serverPath, localPath) {
    let rows = [];
    try {
      rows = await listFolderChildren(siteUrl, library.id, serverPath, getDigest);
    } catch (err) {
      console.error(`[SharePoint Downloader] Falha ao listar "${serverPath}":`, err);
      if (progressCallback) progressCallback({ type: 'error', message: err.message });
      // Propaga o erro para o primeiro nível (pasta raiz) para que o app mostre
      if (serverPath === folderServerPath) throw err;
      return; // Subpastas com erro: pula e continua
    }

    const subfolders = [];
    for (const row of rows) {
      const safeName = sanitizePathSegment(row.FileLeafRef);
      const itemLocalPath = localPath ? `${localPath}/${safeName}` : safeName;

      if (String(row.FSObjType) === '1') {
        if (!SYSTEM_FOLDERS.has(row.FileLeafRef)) subfolders.push({ serverPath: row.FileRef, localPath: itemLocalPath });
        continue;
      }

      files.push({
        name: row.FileLeafRef,
        url: `${origin}${row.FileRef}`,
        localPath: itemLocalPath,
        size: parseInt(row.File_x0020_Size, 10) || 0
      });
    }

    if (progressCallback) progressCallback({ type: 'progress', count: files.length });

    if (!includeSubfolders) return;

    for (const folder of subfolders) {
      await processFolder(folder.serverPath, folder.localPath);
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
