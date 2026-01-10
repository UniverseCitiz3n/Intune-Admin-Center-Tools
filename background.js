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
