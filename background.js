let msGraphToken = null;
const REPORT_REQUESTS_STORAGE_KEY = 'lastCapturedReportRequests';
const REPORT_STATUS_ENDPOINT = 'getDeviceStatusByCompliacePolicyReport';
const REPORT_SUMMARY_ENDPOINT = 'getDeviceStatusSummaryByCompliacePolicyReport';
const isTrustedIntuneRequestSource = (value) => {
  try {
    const { hostname, origin, protocol } = new URL(value);
    if (protocol !== 'https:') return false;

    if (origin === 'https://intune.microsoft.com') {
      return true;
    }

    return hostname.endsWith('.reactblade.portal.azure.net') || hostname.endsWith('.portal.azure.net');
  } catch (error) {
    return false;
  }
};
const hasTrustedRequestSource = (details) => {
  const requestSources = [
    details?.initiator,
    details?.originUrl,
    details?.documentUrl
  ].filter(Boolean);

  return requestSources.some(isTrustedIntuneRequestSource);
};
const shouldCaptureReportRequest = (details) => {
  if (!details || details.method !== 'POST') return false;
  if (!details.url || !details.url.includes('/deviceManagement/reports/')) return false;
  return hasTrustedRequestSource(details);
};

const normalizeReportRequestStore = (store) => {
  if (store && typeof store === 'object' && (store.byTabId || store.byPolicyId)) {
    return {
      byTabId: store.byTabId && typeof store.byTabId === 'object' ? store.byTabId : {},
      byPolicyId: store.byPolicyId && typeof store.byPolicyId === 'object' ? store.byPolicyId : {}
    };
  }

  if (store && typeof store === 'object') {
    return { byTabId: { ...store }, byPolicyId: {} };
  }

  return { byTabId: {}, byPolicyId: {} };
};

const extractPolicyIdFromReportUrl = (url) => {
  const match = (url || '').match(/\/policyId\/([0-9a-f-]{36})(?:\/|$)/i);
  return match ? match[1] : null;
};

const extractPolicyIdsFromRequest = (body, documentUrl = '') => {
  const policyIds = new Set();
  const collectMatches = (text) => {
    if (!text) return;
    const pattern = /PolicyId[\s\S]{0,80}?([0-9a-f-]{36})/ig;
    let match;
    while ((match = pattern.exec(text)) !== null) {
      if (match[1]) {
        policyIds.add(match[1]);
      }
    }
  };

  if (body && typeof body === 'object') {
    collectMatches(JSON.stringify(body));
    if (typeof body.filter === 'string') {
      collectMatches(body.filter);
    }
  } else if (typeof body === 'string') {
    collectMatches(body);
  }

  const policyIdFromUrl = extractPolicyIdFromReportUrl(documentUrl);
  if (policyIdFromUrl) {
    policyIds.add(policyIdFromUrl);
  }

  return [...policyIds];
};

const getReportRequestKind = (url = '') => {
  if (url.includes(REPORT_STATUS_ENDPOINT)) return 'status';
  if (url.includes(REPORT_SUMMARY_ENDPOINT)) return 'summary';
  return 'other';
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
      const existing = normalizeReportRequestStore(data[REPORT_REQUESTS_STORAGE_KEY]);
      const policyIds = extractPolicyIdsFromRequest(parsedBody, details.documentUrl || '');
      const reportRequestEntry = {
        url: details.url,
        method: details.method,
        body: parsedBody,
        capturedAt: new Date(details.timeStamp || Date.now()).toISOString(),
        initiator: requestSource,
        documentUrl: details.documentUrl || null
      };

      if (details.tabId >= 0) {
        existing.byTabId[String(details.tabId)] = reportRequestEntry;
      }

      const requestKind = getReportRequestKind(details.url);
      policyIds.forEach((policyId) => {
        existing.byPolicyId[policyId] = existing.byPolicyId[policyId] || { requests: {} };
        existing.byPolicyId[policyId].lastRequest = reportRequestEntry;
        existing.byPolicyId[policyId].capturedAt = reportRequestEntry.capturedAt;
        existing.byPolicyId[policyId].requests[requestKind] = reportRequestEntry;
      });

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
    extractPolicyIdFromReportUrl,
    extractPolicyIdsFromRequest,
    getReportRequestKind,
    hasTrustedRequestSource,
    isTrustedIntuneRequestSource,
    normalizeReportRequestStore,
    shouldCaptureReportRequest
  };
}
