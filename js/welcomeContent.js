/**
 * Welcome Notification Content Configuration
 * 
 * This file contains the content for welcome notifications for different versions.
 * Update this file when releasing new versions to inform users about changes.
 */

const WELCOME_CONTENT = {
  // Version 1.5.0 - Latest Release
  '1.5.0': {
    title: 'Welcome to Intune Admin Center Tools v1.5.0!',
    intro: `This major release introduces powerful group member management capabilities and enhanced information display for better administration.`,
    changelog: [
      'NEW: Bulk Remove - Remove members from groups (selected or all)',
      'NEW: Enhanced Check Members with exact counts for large groups (1000+)',
      'NEW: Group type indicator (Assigned/Dynamic)',
      'NEW: Dynamic membership rule display in collapsible section',
      'NEW: Row selection support in group members table',
      'IMPROVED: Validation errors now show as auto-dismissing tooltip bubbles',
      'IMPROVED: Better user experience with more informative displays'
    ],
    features: [
      'Search and manage Azure AD groups',
      'Check device/user configuration assignments',
      'Manage group memberships (add/remove devices and users)',
      'Bulk remove (NEW) - Remove selected or all members',
      'Check group members with enhanced details (NEW)',
      'View exact member counts for large groups (NEW)',
      'See group type and dynamic rules (NEW)',
      'Check compliance policy assignments',
      'View app assignments',
      'Check group assignments',
      'Export output tables to CSV',
      'Dark/Light theme support',
      'Download PowerShell scripts for automation'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on any Intune Admin Center page',
      'Switch between Device and User modes using the toggle buttons',
      'Click on rows in the group members table to select specific members',
      'Use the "Bulk Remove" button to remove members from assigned groups',
      'For dynamic groups, expand the membership rule to see the query',
      'Filter results using the search boxes for better navigation',
      'Click the download icon in the pagination area to export the table to CSV'
    ]
  },

  // Previous versions are hidden to keep welcome message focused on latest features
  
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
