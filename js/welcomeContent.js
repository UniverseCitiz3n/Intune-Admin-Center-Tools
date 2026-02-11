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
    intro: `This major release introduces powerful bulk operations, device group creation, and enhanced group member management capabilities for better administration.`,
    changelog: [
      'NEW: Bulk Add - Add multiple members to groups at once (paste lists of emails, UPNs, device names)',
      'NEW: Bulk Remove - Remove members from groups (selected or all)',
      'NEW: Create Device Group from Users - Automatically create device groups based on users\' primary devices',
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
      'Bulk Add (NEW) - Add multiple members at once from pasted lists',
      'Bulk Remove (NEW) - Remove selected or all members',
      'Create Device Group from Users (NEW) - Auto-create device groups',
      'Check group members with enhanced details (NEW)',
      'View exact member counts for large groups (NEW)',
      'See group type and dynamic rules (NEW)',
      'Check compliance policy assignments',
      'View app assignments',
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
  '1.4.1': {
    title: 'Welcome to Intune Admin Center Tools v1.4.1!',
    intro: `This release adds a streamlined way to review configuration assignments for your groups, along with UX improvements.`,
    changelog: [
      'UI: Welcome Guide closing behavior improved',
      'v1.4 features:',
      '  - NEW: Check group assignments feature',
      '  - NEW: Export table data to CSV',
      '  - UI: Freshened user interface for better usability'
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
      'The extension works on any Intune Admin Center page',
      'Switch between Device and User modes using the toggle buttons',
      'Click on rows in the group members table to select specific members',
      'Use "Bulk Add" to paste lists of emails, UPNs, or device names for quick member addition',
      'Use "Bulk Remove" button to remove members from assigned groups',
      'Use "Create Device Group" to build a device group from a user group\'s primary devices',
      'For dynamic groups, expand the membership rule to see the query',
      'Filter results using the search boxes for better navigation',
      'Click the download icon in the pagination area to export the table to CSV'
    ]
  },

  // Previous versions are hidden to keep welcome message focused on latest features
  
  '1.4.2': {
    title: 'Welcome to Intune Admin Center Tools v1.4.2!',
    intro: `This release fixes an important issue with CSV exports for international users.`,
    changelog: [
      'FIX: Added UTF-8 BOM to CSV files for proper encoding in Excel and other applications',
      'v1.4.1 features:',
      '  - UI: Welcome Guide closing behavior improved',
      'v1.4 features:',
      '  - NEW: Check group assignments feature',
      '  - NEW: Export table data to CSV',
      '  - UI: Freshened user interface for better usability'
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
      'Export output tables to CSV with proper character encoding'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on any Intune Admin Center page',
      'Switch between Device and User modes using the toggle buttons',
      'Filter results using the search boxes for better navigation',
      'Click the download icon in the pagination area to export the table to CSV',
      'CSV exports now support UTF-8 with BOM for better compatibility'
    ]
  },

  '1.4.4': {
    title: 'Welcome to Intune Admin Center Tools v1.4.4!',
    intro: `This release adds optional, privacy-focused analytics to help shape future development.`,
    changelog: [
      'NEW: Optional anonymous usage analytics (disabled by default)',
      'Analytics helps understand which features are most valuable',
      'Opt-in via Settings menu: "Enable to help shape the roadmap"',
      'Zero personal data collected - only button clicks and feature usage',
      'Full transparency: see ANALYTICS.md for complete details'
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
      'Export output tables to CSV with proper character encoding',
      'Copy text from tables and search results',
      'Optional analytics toggle (new)'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on any Intune Admin Center page',
      'Switch between Device and User modes using the toggle buttons',
      'Filter results using the search boxes for better navigation',
      'Click the download icon in the pagination area to export the table to CSV',
      'Click and drag to select text in tables and group names, then copy with Ctrl+C',
      'Enable analytics in Settings to help prioritize new features'
    ]
  },

  '1.4.3': {
    title: 'Welcome to Intune Admin Center Tools v1.4.3!',
    intro: `This release enables text selection and copying from tables and group search results.`,
    changelog: [
      'NEW: Text selection enabled in all data tables',
      'NEW: Text selection enabled in group search results',
      'Users can now copy policy names, group names, and other data to clipboard'
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
      'Export output tables to CSV with proper character encoding',
      'Copy text from tables and search results (new)'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on any Intune Admin Center page',
      'Switch between Device and User modes using the toggle buttons',
      'Filter results using the search boxes for better navigation',
      'Click the download icon in the pagination area to export the table to CSV',
      'Click and drag to select text in tables and group names, then copy with Ctrl+C'
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
