// Abre a aba do app ao clicar no ícone da extensão.
// Guarda o tab ID do SharePoint que estava ativo para o app se conectar.

chrome.action.onClicked.addListener(async (tab) => {
  // Se o usuário clicou estando numa aba do SharePoint, salva o ID
  if (tab.url && tab.url.includes('.sharepoint.com')) {
    await chrome.storage.session.set({ lastSpTabId: tab.id });
  }

  // Abre ou foca a aba do app
  const appUrl = chrome.runtime.getURL('app.html');
  const [existing] = await chrome.tabs.query({ url: appUrl });

  if (existing) {
    await chrome.tabs.update(existing.id, { active: true });
    await chrome.windows.update(existing.windowId, { focused: true });
  } else {
    await chrome.tabs.create({ url: appUrl });
  }
});
