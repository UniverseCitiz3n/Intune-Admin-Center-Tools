// Global notification state management
let notificationTimeout = null;
let isProcessing = false;

function showNotification(message, type = 'success', persistent = false) {
  const notificationBar = document.getElementById('notification-bar');
  if (notificationBar) {
    // Clear any existing timeout
    if (notificationTimeout) {
      clearTimeout(notificationTimeout);
      notificationTimeout = null;
    }

    // Update notification content and style
    notificationBar.textContent = message;
    notificationBar.className = type;
    notificationBar.style.display = 'block';

    // Handle auto-dismiss based on type and persistence
    if (persistent) {
      // Persistent notifications stay until replaced or manually dismissed
      isProcessing = true;
    } else {
      isProcessing = false;
      // Auto-dismiss success/error notifications after 2 seconds
      if (type === 'success' || type === 'error') {
        notificationTimeout = setTimeout(() => {
          hideNotification();
        }, 2000);
      } else if (type === 'info' || type === 'warning') {
        // Info and warning notifications dismiss after 3 seconds
        notificationTimeout = setTimeout(() => {
          hideNotification();
        }, 3000);
      }
    }
  } else {
    console.warn('Notification bar not found, falling back to console:', message);
    console.log(message);
  }
}

function showProcessingNotification(message) {
  showNotification(message, 'processing', true);
}

function showResultNotification(message, type = 'success') {
  isProcessing = false;
  showNotification(message, type, false);
}

function hideNotification() {
  const notificationBar = document.getElementById('notification-bar');
  if (notificationBar) {
    notificationBar.style.display = 'none';
    if (notificationTimeout) {
      clearTimeout(notificationTimeout);
      notificationTimeout = null;
    }
    isProcessing = false;
  }
}

function isNotificationProcessing() {
  return isProcessing;
}