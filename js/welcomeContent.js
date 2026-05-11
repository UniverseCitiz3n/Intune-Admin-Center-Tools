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
    intro: `This major release introduces powerful bulk operations, device group creation, Intune device checking, enhanced compliance support, and many more improvements since v1.4.5.`,
    changelog: [
      'NEW: Bulk Add - Add multiple members to groups at once (paste lists of emails, UPNs, device names)',
      'NEW: Bulk Remove - Remove members from groups (selected or all)',
      'NEW: Create Device Group from Users - Automatically create device groups based on users\' primary devices',
      'NEW: Check Intune Devices - See Intune-managed devices for all members of a group',
      'NEW: Not Found list in Check Intune Devices - Expandable list of members with no Intune device found',
      'NEW: Persistent action history for Create Device Group and Check Intune Devices across popup sessions',
      'NEW: CSV export for Check Intune Devices and Create Device Group results',
      'NEW: Selectable column filters in results tables',
      'NEW: Copy auth token to clipboard from the settings menu',
      'NEW: Report Bug shortcut in the settings menu',
      'NEW: Group type indicator (Assigned/Dynamic) with collapsible membership rule display',
      'NEW: Row selection support in group members table',
      'NEW: Enhanced Check Members with exact counts for large groups (1000+)',
      'IMPROVED: Compliance assignment check now includes Settings Catalog policies and correctly scopes results by device platform',
      'IMPROVED: Log collection now works for shared devices without a primary user',
      'IMPROVED: Dynamic membership rule and group type are now persisted across popup reloads',
      'IMPROVED: Group selection correctly takes precedence over cached group for all bulk actions',
      'IMPROVED: Validation errors now show as auto-dismissing tooltip bubbles',
      'FIXED: Support for new Intune preview device view URL structure (managedDeviceId segment)'
    ],
    features: [
      'Search and manage Azure AD groups',
      'Check device/user configuration assignments',
      'Manage group memberships (add/remove devices and users)',
      'Bulk Add - Add multiple members at once from pasted lists',
      'Bulk Remove - Remove selected or all members',
      'Create Device Group from Users - Auto-create device groups with persistent history',
      'Check Intune Devices - View managed devices for all group members',
      'Check group members with enhanced details and exact counts for large groups (1000+)',
      'See group type (Assigned/Dynamic) and dynamic membership rules',
      'Selectable column filters in all results tables',
      'Copy auth token to clipboard from the settings menu',
      'Check compliance policy assignments (including Settings Catalog)',
      'View app assignments',
      'Check group assignments',
      'Export output tables to CSV',
      'Report Bug shortcut in the settings menu'
    ],
    tips: [
      'Use Ctrl+Shift+W to show this welcome message anytime',
      'The extension works on any Intune Admin Center page',
      'Switch between Device and User modes using the toggle buttons',
      'Click on rows in the group members table to select specific members for bulk actions',
      'Use "Bulk Add" to paste lists of emails, UPNs, or device names for quick member addition',
      'Use "Bulk Remove" to remove selected or all members from a group',
      'Use "Create Device Group" to build a device group from a user group\'s primary devices',
      'Use "Check Intune Devices" to see which Intune-managed devices belong to group members',
      'For dynamic groups, expand the membership rule to see the query',
      'Use the column filter button in results tables to show only the columns you need',
      'Copy the captured auth token from Settings > Copy auth token for use in API calls or scripts',
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
