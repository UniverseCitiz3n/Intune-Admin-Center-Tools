let msGraphToken = null;
const REPORT_REQUESTS_STORAGE_KEY = 'lastCapturedReportRequests';
const isTrustedIntuneRequestSource = (value) => {
  try {
    return new URL(value).origin === 'https://intune.microsoft.com';
  } catch (error) {
    return false;
  }
};
const shouldCaptureReportRequest = (details) => {
  if (!details || details.tabId < 0 || details.method !== 'POST') return false;
  if (!details.url || !details.url.includes('/deviceManagement/reports/')) return false;
  const requestSource = details.initiator || details.originUrl || details.documentUrl || '';
  return isTrustedIntuneRequestSource(requestSource);
};

const decodeRequestBody = (requestBody) => {
  if (!requestBody) return null;

  if (requestBody.raw && requestBody.raw.length > 0) {
    const combinedLength = requestBody.raw.reduce((total, part) => total + (part.bytes ? part.bytes.byteLength : 0), 0);
    const combined = new Uint8Array(combinedLength);
    let offset = 0;

    requestBody.raw.forEach(part => {
      if (!part.bytes) return;
      const bytes = new Uint8Array(part.bytes);
      combined.set(bytes, offset);
      offset += bytes.byteLength;
    });

    return new TextDecoder().decode(combined);
  }

  if (requestBody.formData) {
    return JSON.stringify(requestBody.formData);
  }

  return null;
};

const persistReportRequest = (details) => {
  if (!shouldCaptureReportRequest(details)) return;
  const requestSource = details.initiator || details.originUrl || details.documentUrl || '';

  const requestBody = decodeRequestBody(details.requestBody);
  if (!requestBody) return;

  let parsedBody;
  try {
    parsedBody = JSON.parse(requestBody);
  } catch (error) {
    console.log('Skipping non-JSON report request body:', error.message);
    return;
  }

  if (typeof chrome !== 'undefined' && chrome.storage?.local) {
    chrome.storage.local.get([REPORT_REQUESTS_STORAGE_KEY], (data) => {
      const existing = data[REPORT_REQUESTS_STORAGE_KEY] || {};
      existing[String(details.tabId)] = {
        url: details.url,
        method: details.method,
        body: parsedBody,
        capturedAt: new Date(details.timeStamp || Date.now()).toISOString(),
        initiator: requestSource,
        documentUrl: details.documentUrl || null
        };
      chrome.storage.local.set({ [REPORT_REQUESTS_STORAGE_KEY]: existing });
    });
  }
};

if (typeof chrome !== 'undefined') {
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

  chrome.webRequest.onBeforeRequest.addListener(
    (details) => {
      persistReportRequest(details);
    },
    { urls: ["*://graph.microsoft.com/*"] },
    ["requestBody"]
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
      // Set default analytics preference on installation
      // For beta releases, analytics is enabled by default
      // For prod releases, analytics is disabled by default
      // Note: The actual release type check is in analytics.js
      // We set to undefined here to let analytics.js handle the default
      // This ensures consistent behavior between first install and updates
    } else if (details.reason === 'update') {
      console.log('[Analytics] Extension updated from', details.previousVersion, 'to', chrome.runtime.getManifest().version);
    }
  });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    decodeRequestBody,
    isTrustedIntuneRequestSource,
    shouldCaptureReportRequest
  };
}
