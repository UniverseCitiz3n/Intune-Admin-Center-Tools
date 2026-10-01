const test = require('node:test');
const assert = require('node:assert/strict');

const {
  decodeRequestBody,
  extractPolicyIdFromReportUrl,
  extractPolicyIdsFromRequest,
  getReportRequestKind,
  hasTrustedRequestSource,
  isTrustedIntuneRequestSource,
  normalizeReportRequestStore,
  shouldCaptureReportRequest
} = require('./background.js');

test('accepts only the Intune origin for report capture', () => {
  assert.equal(isTrustedIntuneRequestSource('https://intune.microsoft.com/#view/foo'), true);
  assert.equal(isTrustedIntuneRequestSource('https://sandbox-2.reactblade.portal.azure.net/#view/foo'), true);
  assert.equal(isTrustedIntuneRequestSource('https://main.portal.azure.net/#view/foo'), true);
  assert.equal(isTrustedIntuneRequestSource('https://intune.microsoft.com.evil.example/#view/foo'), false);
  assert.equal(isTrustedIntuneRequestSource('not-a-url'), false);
});

test('accepts any trusted source field on report requests', () => {
  assert.equal(hasTrustedRequestSource({
    initiator: 'https://sandbox-2.reactblade.portal.azure.net'
  }), true);
  assert.equal(hasTrustedRequestSource({
    originUrl: 'https://main.portal.azure.net'
  }), true);
  assert.equal(hasTrustedRequestSource({
    documentUrl: 'https://intune.microsoft.com/#view/foo'
  }), true);
  assert.equal(hasTrustedRequestSource({
    initiator: 'https://example.com',
    documentUrl: 'https://evil.example'
  }), false);
});

test('identifies capture-worthy Intune report requests', () => {
  assert.equal(shouldCaptureReportRequest({
    tabId: 5,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    initiator: 'https://sandbox-2.reactblade.portal.azure.net'
  }), true);

  assert.equal(shouldCaptureReportRequest({
    tabId: -1,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    initiator: 'https://intune.microsoft.com'
  }), true);

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

  assert.equal(shouldCaptureReportRequest({
    tabId: 5,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    originUrl: 'https://intune.microsoft.com'
  }), true);

  assert.equal(shouldCaptureReportRequest({
    tabId: 5,
    method: 'POST',
    url: 'https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusByCompliacePolicyReport',
    documentUrl: 'https://intune.microsoft.com/#view/foo'
  }), true);
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

test('decodes multipart raw bodies and form data payloads', () => {
  const multipartBody = {
    raw: [
      { bytes: new TextEncoder().encode('{"filter":"Policy').buffer },
      { bytes: new TextEncoder().encode('Status eq 4"}').buffer }
    ]
  };
  const formDataBody = {
    formData: {
      filter: ['PolicyStatus eq 4'],
      top: ['50']
    }
  };

  assert.equal(decodeRequestBody(multipartBody), '{"filter":"PolicyStatus eq 4"}');
  assert.equal(decodeRequestBody(formDataBody), JSON.stringify(formDataBody.formData));
});

test('extracts policy IDs and normalizes report storage', () => {
  assert.equal(
    extractPolicyIdFromReportUrl('https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DeviceStatusByCompliacePolicy.ReactView/policyId/7ca45552-b08b-47df-9b51-8a40abbcfb50/foo'),
    '7ca45552-b08b-47df-9b51-8a40abbcfb50'
  );

  assert.deepEqual(
    extractPolicyIdsFromRequest({
      filter: "(PolicyId eq '7ca45552-b08b-47df-9b51-8a40abbcfb50') and (PolicyStatus eq '4')"
    }),
    ['7ca45552-b08b-47df-9b51-8a40abbcfb50']
  );

  assert.equal(
    getReportRequestKind('https://graph.microsoft.com/beta/deviceManagement/reports/getDeviceStatusSummaryByCompliacePolicyReport'),
    'summary'
  );

  assert.deepEqual(
    normalizeReportRequestStore({ legacy: { body: {} } }),
    { byTabId: { legacy: { body: {} } }, byPolicyId: {} }
  );
});
