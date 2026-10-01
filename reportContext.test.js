const test = require('node:test');
const assert = require('node:assert/strict');

const {
  isDeviceContextUrl,
  normalizeIntunePageContext,
  isKnownReportAddContextUrl,
  extractPolicyIdFromReportUrl,
  getBestCapturedReportRequest,
  buildStatusReportRequestFromCaptured,
  isSupportedReportAddContext
} = require('./js/reportContext.js');

test('identifies device and report add contexts', () => {
  assert.equal(isDeviceContextUrl('https://intune.microsoft.com/#view/.../managedDeviceId/1234'), true);
  assert.equal(
    isKnownReportAddContextUrl('https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DeviceStatusByCompliacePolicy.ReactView/policyId/abc'),
    true
  );
  assert.equal(
    isKnownReportAddContextUrl('https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/overview'),
    false
  );
});

test('normalizes Intune hash contexts for comparisons', () => {
  assert.equal(
    normalizeIntunePageContext('https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DeviceStatusByCompliacePolicy.ReactView/policyId/abc?foo=bar'),
    'https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DeviceStatusByCompliacePolicy.ReactView/policyId/abc'
  );
});

test('extracts policy IDs from report URLs', () => {
  assert.equal(
    extractPolicyIdFromReportUrl('https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DeviceStatusByCompliacePolicy.ReactView/policyId/3ee52e60-bbc8-4a83-aa32-a4e35ee2a62a/policyPlatformType~/6'),
    '3ee52e60-bbc8-4a83-aa32-a4e35ee2a62a'
  );
});

test('supports report replay on known report pages and preserves device-page fallback', () => {
  const reportUrl = 'https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DeviceStatusByCompliacePolicy.ReactView/policyId/abc';

  assert.equal(
    isSupportedReportAddContext(reportUrl, {
      url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
      documentUrl: 'https://intune.microsoft.com/'
    }),
    true
  );

  assert.equal(
    isSupportedReportAddContext('https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/managedDeviceId/1234', {
      url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport'
    }),
    false
  );

  assert.equal(isSupportedReportAddContext(reportUrl, null), false);
});

test('prefers the latest policy request and preserves filter/search when converting to status report', () => {
  const reportUrl = 'https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DeviceStatusByCompliacePolicy.ReactView/policyId/3ee52e60-bbc8-4a83-aa32-a4e35ee2a62a/policyPlatformType~/6';
  const storedRequests = {
    byTabId: {
      '7': {
        url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
        method: 'POST',
        body: { filter: "(PolicyId eq '3ee52e60-bbc8-4a83-aa32-a4e35ee2a62a')", top: 50 },
        capturedAt: '2026-10-01T11:00:00.000Z'
      }
    },
    byPolicyId: {
      '3ee52e60-bbc8-4a83-aa32-a4e35ee2a62a': {
        lastRequest: {
          url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusSummaryByCompliacePolicyReport',
          method: 'POST',
          body: {
            filter: "(PolicyId eq '3ee52e60-bbc8-4a83-aa32-a4e35ee2a62a') and ((OS eq 'WindowsHolographic') or (OS eq 'Windows')) and substringof('HQ', DeviceName)",
            search: 'HQ'
          },
          capturedAt: '2026-10-01T11:05:00.000Z',
          documentUrl: reportUrl
        }
      }
    }
  };

  const bestRequest = getBestCapturedReportRequest(storedRequests, reportUrl, 7, new Date('2026-10-01T11:06:00.000Z').getTime());
  const statusRequest = buildStatusReportRequestFromCaptured(bestRequest, reportUrl);

  assert.equal(bestRequest.url.includes('getDeviceStatusSummaryByCompliacePolicyReport'), true);
  assert.equal(statusRequest.url.includes('getDeviceStatusByCompliacePolicyReport'), true);
  assert.equal(statusRequest.body.filter.includes("substringof('HQ', DeviceName)"), true);
  assert.equal(statusRequest.body.search, 'HQ');
});
