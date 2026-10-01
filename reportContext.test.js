const test = require('node:test');
const assert = require('node:assert/strict');

const {
  isDeviceContextUrl,
  normalizeIntunePageContext,
  isKnownReportAddContextUrl,
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
