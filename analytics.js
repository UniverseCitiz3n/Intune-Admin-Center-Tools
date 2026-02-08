// analytics.js - Privacy-focused Google Analytics for Intune Admin Center Tools
// Uses GA4 Measurement Protocol API to track usage without collecting PII
//
// SETUP INSTRUCTIONS:
// 1. Create a Google Analytics 4 property at https://analytics.google.com
// 2. Get your Measurement ID (format: G-XXXXXXXXXX) from Admin > Data Streams
// 3. Create a Measurement Protocol API Secret at Admin > Data Streams > [your stream] > Measurement Protocol API secrets
// 4. Replace GA_MEASUREMENT_ID and GA_API_SECRET below with your actual values
// 5. Deploy the extension
//
// SECURITY NOTE:
// The credentials below are PLACEHOLDERS and will not work until replaced.
// For production use, consider using environment variables during build time
// or a configuration file that is not committed to the repository.
//
// PRIVACY NOTICE:
// This implementation tracks ONLY:
// - Button clicks and feature usage (no content)
// - Anonymous client ID (randomly generated UUID)
// - Extension version
// - Session information
// NO personal information, device IDs, user names, tokens, or Microsoft Graph data is collected.

const Analytics = (() => {
  // GA4 Configuration - REPLACE THESE VALUES WITH YOUR ACTUAL CREDENTIALS
  // These are PLACEHOLDERS and will not send data until configured
  const GA_MEASUREMENT_ID = 'G-XXXXXXXXXX'; // Replace with actual GA4 Measurement ID
  const GA_API_SECRET = 'XXXXXXXXXXXXXXXXXXXX'; // Replace with actual API Secret
  const MEASUREMENT_PROTOCOL_URL = `https://www.google-analytics.com/mp/collect?measurement_id=${GA_MEASUREMENT_ID}&api_secret=${GA_API_SECRET}`;
  
  let clientId = null;
  let analyticsEnabled = true; // Default to enabled, user can opt-out
  let sessionId = null;

  // Initialize analytics
  const init = async () => {
    // Load settings from chrome.storage
    const data = await chrome.storage.local.get(['analyticsEnabled', 'analyticsClientId', 'analyticsSessionId']);
    
    // Check if user has opted out
    if (data.analyticsEnabled === false) {
      analyticsEnabled = false;
      console.log('[Analytics] User has opted out of analytics');
      return;
    }
    
    // Generate or retrieve client ID (anonymous, persistent identifier)
    if (data.analyticsClientId) {
      clientId = data.analyticsClientId;
    } else {
      clientId = generateClientId();
      await chrome.storage.local.set({ analyticsClientId: clientId });
    }
    
    // Generate session ID (temporary, per-session identifier)
    sessionId = Date.now().toString();
    await chrome.storage.local.set({ analyticsSessionId: sessionId });
    
    console.log('[Analytics] Initialized with client ID:', clientId);
  };

  // Generate anonymous client ID
  const generateClientId = () => {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = Math.random() * 16 | 0;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  };

  // Send event to Google Analytics
  const sendEvent = async (eventName, eventParams = {}) => {
    if (!analyticsEnabled || !clientId) {
      return;
    }

    // Get extension version from manifest
    const manifest = chrome.runtime.getManifest();
    const extensionVersion = manifest.version;

    // Prepare GA4 Measurement Protocol payload
    const payload = {
      client_id: clientId,
      events: [{
        name: eventName,
        params: {
          session_id: sessionId,
          engagement_time_msec: '100',
          extension_version: extensionVersion,
          ...eventParams
        }
      }]
    };

    try {
      await fetch(MEASUREMENT_PROTOCOL_URL, {
        method: 'POST',
        body: JSON.stringify(payload)
      });
      console.log('[Analytics] Event sent:', eventName, eventParams);
    } catch (error) {
      console.error('[Analytics] Failed to send event:', error);
    }
  };

  // Track button click
  const trackButtonClick = (buttonName) => {
    sendEvent('button_click', {
      button_name: buttonName
    });
  };

  // Track feature usage
  const trackFeatureUsage = (featureName, additionalParams = {}) => {
    sendEvent('feature_usage', {
      feature_name: featureName,
      ...additionalParams
    });
  };

  // Track page view (popup open)
  const trackPageView = (pageName = 'popup') => {
    sendEvent('page_view', {
      page_title: pageName,
      page_location: pageName
    });
  };

  // Track error
  const trackError = (errorType, errorMessage) => {
    sendEvent('error', {
      error_type: errorType,
      error_message: errorMessage.substring(0, 100) // Limit to 100 chars to avoid PII
    });
  };

  // Enable analytics
  const enable = async () => {
    analyticsEnabled = true;
    await chrome.storage.local.set({ analyticsEnabled: true });
    
    // Initialize if not already done
    if (!clientId) {
      await init();
    }
    
    console.log('[Analytics] Enabled');
    sendEvent('analytics_enabled');
  };

  // Disable analytics
  const disable = async () => {
    // Send event before disabling
    if (analyticsEnabled && clientId) {
      await sendEvent('analytics_disabled');
    }
    analyticsEnabled = false;
    await chrome.storage.local.set({ analyticsEnabled: false });
    console.log('[Analytics] Disabled');
  };

  // Check if analytics is enabled
  const isEnabled = () => {
    return analyticsEnabled;
  };

  // Generic event tracking
  const trackEvent = (eventName, params = {}) => {
    sendEvent(eventName, params);
  };

  return {
    init,
    trackButtonClick,
    trackFeatureUsage,
    trackPageView,
    trackError,
    trackEvent,
    enable,
    disable,
    isEnabled
  };
})();

// Make Analytics available globally
if (typeof window !== 'undefined') {
  window.Analytics = Analytics;
}
