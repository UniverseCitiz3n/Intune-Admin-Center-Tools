function logMessage(msg) {
  if (typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.sendMessage) {
    chrome.runtime.sendMessage({ type: "LOG_MESSAGE", payload: msg });
  } else {
    console.log(msg);
  }
}