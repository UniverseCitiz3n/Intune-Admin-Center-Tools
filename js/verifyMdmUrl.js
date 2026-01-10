function verifyMdmUrl() {
  return new Promise((resolve, reject) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (!tabs || !tabs[0]) {
        const error = 'No active tab found.';
        logMessage(error);
        showNotification(error, 'error');
        return reject(new Error(error));
      }
      const url = tabs[0].url;
      const sanitizedUrl = url.replace(/\/[\w-]+(?=\/|$)/g, '/[REDACTED]');
      logMessage(`Active tab URL: ${sanitizedUrl}`);
      const mdmMatch = url.match(/mdmDeviceId\/([\w-]+)/i);
      if (!mdmMatch) {
        const error = 'mdmDeviceId not found in URL.';
        logMessage(error);
        showNotification(error, 'error');
        return reject(new Error(error));
      }
      resolve({ mdmDeviceId: mdmMatch[1] });
    });
  });
}