(function (global) {
  const REPORT_ADD_URL_PATTERNS = [
    /Microsoft_Intune_DeviceSettings\/DeviceStatusByCompliacePolicy\.ReactView/i
  ];

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
    isSupportedReportAddContext
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = helpers;
  }

  global.reportContextHelpers = helpers;
})(typeof globalThis !== 'undefined' ? globalThis : this);
