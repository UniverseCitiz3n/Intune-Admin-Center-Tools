# Google Analytics Integration

## Overview

This extension includes optional, privacy-focused Google Analytics to help understand how features are used. The analytics implementation is designed with **privacy first** - no personal, sensitive, or Microsoft-related data is collected.

## What Data is Collected

The analytics system tracks **only** the following information:

### Tracked Events
- **Button Clicks**: Which buttons are clicked (e.g., "Check Configuration Assignments", "Search Groups")
- **Feature Usage**: Which features are used (e.g., CSV export, theme toggle, target mode switch)
- **Page Views**: When the extension popup is opened
- **Extension Version**: The version number of the installed extension

### Anonymous Identifiers
- **Client ID**: A randomly generated UUID that persists across sessions
- **Session ID**: A temporary identifier for the current browsing session

## What is NOT Collected

The following information is **NEVER** collected or sent to Google Analytics:

❌ Device IDs or serial numbers  
❌ User names or UPNs  
❌ Microsoft Graph tokens  
❌ Group names or IDs  
❌ Policy names or content  
❌ App names or assignments  
❌ PowerShell script content  
❌ Any Microsoft Intune data  
❌ IP addresses (anonymized by GA4)  
❌ Exact timestamps (aggregated by GA4)  

## User Control

### Opt-In
Analytics is **disabled by default** on installation. Users can enable it to help shape the roadmap:

1. Click the **Settings** button (⋮) in the extension popup
2. Click **Analytics: Disabled** to toggle it on
3. A notification will confirm that analytics is enabled

### Opt-Out
Users can disable analytics at any time:

1. Click the **Settings** button (⋮) in the extension popup
2. Click **Analytics: Enabled** to toggle it off
3. A notification will confirm that analytics is disabled

### Checking Status
The current analytics status is always visible in the Settings menu.

## Technical Implementation

### Google Analytics 4 (GA4)
This extension uses GA4 with the Measurement Protocol API, which:
- Does not require loading external JavaScript libraries
- Works with Manifest V3 Content Security Policy
- Sends data directly to Google Analytics servers via HTTPS POST

### Architecture
- **analytics.js**: Core analytics module with privacy-safe event tracking
- **popup.js**: Initializes analytics and tracks user interactions
- **background.js**: Tracks installation and update events
- **manifest.json**: Declares required permissions for Google Analytics domain

### Data Flow
1. User interacts with the extension (clicks button, opens popup, etc.)
2. Analytics module checks if analytics is enabled
3. If enabled, sends anonymous event to GA4 via Measurement Protocol
4. Google Analytics aggregates and reports the data

## Events Reference

### Button Click Events
Event name: `button_click`

Parameters:
- `button_name`: Name of the clicked button
  - `search_group`
  - `add_to_groups`
  - `remove_from_groups`
  - `check_groups`
  - `check_group_members`
  - `check_group_assignments`
  - `check_compliance`
  - `download_script`
  - `apps_assignment`
  - `pwsh_profiles`
  - `collect_logs`
  - `create_group`
  - `theme_toggle`

### Feature Usage Events
Event name: `feature_usage`

Parameters:
- `feature_name`: Name of the feature used
  - `export_csv` (with `display_type`: `config`, `apps`, `compliance`, `scripts`, `groupMembers`, `groupAssignments`)
  - `toggle_target_mode` (with `mode`: `device` or `user`)

### Page View Events
Event name: `page_view`

Parameters:
- `page_title`: Always `popup`
- `page_location`: Always `popup`

### System Events
- `analytics_enabled`: User enabled analytics
- `analytics_disabled`: User disabled analytics (sent before disabling)

## Privacy Compliance

This analytics implementation is designed to comply with:
- **GDPR**: No personal data collected, user consent respected
- **CCPA**: Anonymous data only, opt-out available
- **Browser Extension Policies**: No tracking of browsing behavior outside the extension

## Questions & Concerns

If you have privacy concerns or questions about the analytics:

1. **Review the Code**: All analytics code is in `analytics.js` and is open source
2. **Opt-Out**: You can always disable analytics in Settings
3. **Open an Issue**: Report concerns on [GitHub Issues](../../issues)

## Transparency Statement

This analytics system was added to help the developer understand:
- Which features are most useful
- Which features are rarely used (candidates for removal)
- What buttons users click most frequently
- How often the extension is used

The goal is to improve the extension based on actual usage patterns, not to collect personal information.

**Your privacy is paramount. If you're uncomfortable with analytics, please disable it in Settings.**
