/**
 * Welcome Notification Content Configuration
 * 
 * This file contains the content for welcome notifications for different versions.
 * Update this file when releasing new versions to inform users about changes.
 */

const WELCOME_CONTENT = {
  // Version 1.3 - Current version
  '1.3': {
    title: 'Welcome to Intune Admin Center Tools v1.3!',
    intro: `Thank you for using Intune Admin Center Tools! This extension helps you efficiently manage Intune devices, check assignments, and handle group memberships directly from the Intune Admin Center.`,
    changelog: [
      'NEW: Check group members functionality',
      'Improved user/device targeting',
      'Enhanced group management tools',
      'Better pagination and filtering',
      'Various bug fixes and performance improvements'
    ],
    features: [
      'Search and manage Azure AD groups',
      'Check device/user configuration assignments',
      'Manage group memberships (add/remove devices and users)',
      'Check compliance policy assignments',
      'View app assignments',
      'Dark/Light theme support',
      'Download PowerShell scripts for automation'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on Intune device pages',
      'Switch between Device and User modes using the toggle buttons',
      'Filter results using the search boxes for better navigation'
    ]
  },

  // Version 1.4 - New features
  '1.4': {
    title: 'Welcome to Intune Admin Center Tools v1.4!',
    intro: `This release adds a streamlined way to review configuration assignments for your groups, along with UX improvements.`,
    changelog: [
      'NEW: Check group assignments feature',
      'NEW: Export table data to CSV',
      'UI: Improved button text wrapping to prevent truncation'
    ],
    features: [
      'Search and manage Azure AD groups',
      'Check device/user configuration assignments',
      'Manage group memberships (add/remove devices and users)',
      'Check compliance policy assignments',
      'View app assignments',
      'Dark/Light theme support',
      'Download PowerShell scripts for automation',
      'Check group assignments',
      'Export output tables to CSV (new)'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on any Intune Admin Center page',
      'Switch between Device and User modes using the toggle buttons',
      'Filter results using the search boxes for better navigation',
      'Click the download icon in the pagination area to export the table to CSV'
    ]
  },

  // Default fallback content
  'default': {
    title: 'Welcome to Intune Admin Center Tools!',
    intro: `Thank you for using Intune Admin Center Tools! This extension helps you efficiently manage Intune devices and assignments.`,
    changelog: [
      'General improvements and bug fixes'
    ],
    features: [
      'Search and manage Azure AD groups',
      'Check device/user assignments',
      'Manage group memberships',
      'Dark/Light theme support'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on any Intune Admin Center page'
    ]
  }
};

// Export for use in welcomeNotification.js
if (typeof module !== 'undefined' && module.exports) {
  module.exports = WELCOME_CONTENT;
} else {
  window.WELCOME_CONTENT = WELCOME_CONTENT;
}
