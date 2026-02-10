let msGraphToken = null;

chrome.webRequest.onBeforeSendHeaders.addListener(
  (details) => {
    for (const header of details.requestHeaders) {
      if (header.name.toLowerCase() === 'authorization') {
        msGraphToken = header.value;
        console.log("Token captured:", msGraphToken);
        // Store token in chrome.storage.local so popup.js can access it.
        chrome.storage.local.set({ msGraphToken });
        break;
      }
    }
  },
  { urls: ["*://graph.microsoft.com/*"] },
  ["requestHeaders", "extraHeaders"]
);
// background.js
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "LOG_MESSAGE") {
    console.log("Received from popup:", message.payload);
  }
});

// Track extension installation and updates
chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    console.log('[Analytics] Extension installed');
    // Set default analytics preference on installation (disabled by default)
    chrome.storage.local.set({ analyticsEnabled: false });
  } else if (details.reason === 'update') {
    console.log('[Analytics] Extension updated from', details.previousVersion, 'to', chrome.runtime.getManifest().version);
  }
});
