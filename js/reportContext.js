(function (global) {
  const REPORT_ADD_URL_PATTERNS = [
    /Microsoft_Intune_DeviceSettings\/DeviceStatusByCompliacePolicy\.ReactView/i
  ];
  const REPORT_REQUEST_MAX_AGE_MS = 30 * 60 * 1000;
  const REPORT_STATUS_ENDPOINT = 'getDeviceStatusByCompliacePolicyReport';

  const isDeviceContextUrl = (url) => /(?:mdmDeviceId|managedDeviceId)\//i.test(url || '');

  const normalizeIntunePageContext = (url) => {
    try {
      const parsed = new URL(url);
      const normalizedHash = (parsed.hash || '').split('?')[0].replace(/\/+$/, '');
      return `${parsed.origin}${parsed.pathname}${normalizedHash}`;
    } catch (error) {
      return url || '';
    }
  };

  const isKnownReportAddContextUrl = (url) => {
    if (!url) return false;
    return REPORT_ADD_URL_PATTERNS.some((pattern) => pattern.test(url));
  };

  const extractPolicyIdFromReportUrl = (url) => {
    const match = (url || '').match(/\/policyId\/([0-9a-f-]{36})(?:\/|$)/i);
    return match ? match[1] : null;
  };

  const normalizeStoredReportRequests = (storedRequests) => (
    (storedRequests && (storedRequests.byTabId || storedRequests.byPolicyId))
      ? storedRequests
      : { byTabId: storedRequests || {}, byPolicyId: {} }
  );

  const isFreshReportRequest = (reportRequest, nowMs = Date.now(), maxAgeMs = REPORT_REQUEST_MAX_AGE_MS) => {
    if (!reportRequest || !reportRequest.capturedAt) return false;
    const capturedAt = new Date(reportRequest.capturedAt).getTime();
    return Boolean(capturedAt) && (nowMs - capturedAt) <= maxAgeMs;
  };

  const getBestCapturedReportRequest = (storedRequests, activeTabUrl, tabId, nowMs = Date.now(), maxAgeMs = REPORT_REQUEST_MAX_AGE_MS) => {
    const requests = normalizeStoredReportRequests(storedRequests);
    const currentPolicyId = extractPolicyIdFromReportUrl(activeTabUrl);
    const tabRequest = (typeof tabId === 'number' && requests.byTabId)
      ? requests.byTabId[String(tabId)]
      : null;
    const latestPolicyRequest = currentPolicyId && requests.byPolicyId
      ? requests.byPolicyId[currentPolicyId]?.lastRequest || null
      : null;

    if (isFreshReportRequest(tabRequest, nowMs, maxAgeMs)) {
      if (isFreshReportRequest(latestPolicyRequest, nowMs, maxAgeMs)) {
        const tabCapturedAt = new Date(tabRequest.capturedAt).getTime();
        const policyCapturedAt = new Date(latestPolicyRequest.capturedAt).getTime();
        return policyCapturedAt >= tabCapturedAt ? latestPolicyRequest : tabRequest;
      }

      return tabRequest;
    }

    if (isFreshReportRequest(latestPolicyRequest, nowMs, maxAgeMs)) {
      return latestPolicyRequest;
    }

    return null;
  };

  const buildStatusReportRequestFromCaptured = (reportRequest, activeTabUrl = '') => {
    if (!reportRequest) return null;

    const requestBody = JSON.parse(JSON.stringify(reportRequest.body || {}));
    const policyId = extractPolicyIdFromReportUrl(activeTabUrl || reportRequest.documentUrl || '');
    if (!requestBody.filter && policyId) {
      requestBody.filter = `(PolicyId eq '${policyId}')`;
    }

    if (!Array.isArray(requestBody.select)) {
      requestBody.select = [];
    }

    if (!Array.isArray(requestBody.orderBy) || requestBody.orderBy.length === 0) {
      requestBody.orderBy = ['DeviceName asc'];
    }

    if (requestBody.search === undefined || requestBody.search === null) {
      requestBody.search = '';
    }

    if (requestBody.skip === undefined || requestBody.skip === null) {
      requestBody.skip = 0;
    }

    if (requestBody.top === undefined || requestBody.top === null) {
      requestBody.top = 50;
    }

    return {
      ...reportRequest,
      url: `https://graph.microsoft.com/beta/deviceManagement/reports/${REPORT_STATUS_ENDPOINT}`,
      method: 'POST',
      body: requestBody,
      documentUrl: reportRequest.documentUrl || activeTabUrl || null
    };
  };

  const isSupportedReportAddContext = (activeTabUrl, reportRequest) => {
    if (!reportRequest || isDeviceContextUrl(activeTabUrl)) {
      return false;
    }

    return isKnownReportAddContextUrl(activeTabUrl);
  };

  const helpers = {
    isDeviceContextUrl,
    normalizeIntunePageContext,
    isKnownReportAddContextUrl,
    extractPolicyIdFromReportUrl,
    normalizeStoredReportRequests,
    isFreshReportRequest,
    getBestCapturedReportRequest,
    buildStatusReportRequestFromCaptured,
    isSupportedReportAddContext
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = helpers;
  }

  global.reportContextHelpers = helpers;
})(typeof globalThis !== 'undefined' ? globalThis : this);
