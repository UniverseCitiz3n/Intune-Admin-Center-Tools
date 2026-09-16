const test = require('node:test');
const assert = require('node:assert/strict');

const {
  decodeRequestBody,
  isTrustedIntuneRequestSource,
  shouldCaptureReportRequest
} = require('./background.js');

test('accepts only the Intune origin for report capture', () => {
  assert.equal(isTrustedIntuneRequestSource('https://intune.microsoft.com/#view/foo'), true);
  assert.equal(isTrustedIntuneRequestSource('https://intune.microsoft.com.evil.example/#view/foo'), false);
  assert.equal(isTrustedIntuneRequestSource('not-a-url'), false);
});

test('identifies capture-worthy Intune report requests', () => {
  assert.equal(shouldCaptureReportRequest({
    tabId: 5,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    initiator: 'https://intune.microsoft.com'
  }), true);

  assert.equal(shouldCaptureReportRequest({
    tabId: -1,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    initiator: 'https://intune.microsoft.com'
  }), false);

  assert.equal(shouldCaptureReportRequest({
    tabId: 5,
    method: 'GET',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    initiator: 'https://intune.microsoft.com'
  }), false);

  assert.equal(shouldCaptureReportRequest({
    tabId: 5,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/managedDevices',
    initiator: 'https://intune.microsoft.com'
  }), false);

  assert.equal(shouldCaptureReportRequest({
    tabId: 5,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    initiator: 'https://intune.microsoft.com.evil.example'
  }), false);
});

test('decodes raw JSON request bodies', () => {
  const requestBody = {
    raw: [
      {
        bytes: new TextEncoder().encode('{"filter":"PolicyStatus eq 4"}').buffer
      }
    ]
  };

  assert.equal(decodeRequestBody(requestBody), '{"filter":"PolicyStatus eq 4"}');
});
