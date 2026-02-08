# Google Analytics Configuration Guide

This guide explains how to configure Google Analytics 4 (GA4) for the Intune Admin Center Tools extension.

## Prerequisites

- Google account with access to Google Analytics
- Admin access to the extension repository (for updating configuration)

## Step 1: Create Google Analytics 4 Property

1. Go to [Google Analytics](https://analytics.google.com/)
2. Click **Admin** (gear icon in bottom left)
3. Click **Create Property**
4. Fill in property details:
   - **Property name**: `Intune Admin Center Tools` (or your preferred name)
   - **Reporting time zone**: Select your time zone
   - **Currency**: Select your currency
5. Click **Next**
6. Fill in business details (optional but recommended)
7. Click **Create**
8. Accept the Terms of Service

## Step 2: Create Data Stream

1. After creating the property, you'll be prompted to set up a data stream
2. Select **Web** as the platform type (even though this is a browser extension)
3. Fill in stream details:
   - **Website URL**: `chrome-extension://` (placeholder, not used)
   - **Stream name**: `Intune Admin Center Tools`
4. **Important**: Enable **Enhanced measurement** (toggle on)
5. Click **Create stream**

## Step 3: Get Measurement ID

1. After creating the stream, you'll see the **Web stream details** page
2. Copy your **Measurement ID** at the top (format: `G-XXXXXXXXXX`)
3. Save this for later

## Step 4: Create Measurement Protocol API Secret

1. On the **Web stream details** page, scroll down
2. Find the section **Measurement Protocol API secrets**
3. Click **Create**
4. Fill in secret details:
   - **Nickname**: `Extension API Secret` (or your preferred name)
5. Click **Create**
6. Copy the **Secret value** shown (format: long alphanumeric string)
7. **IMPORTANT**: Save this value immediately - you won't be able to see it again!

## Step 5: Update Extension Configuration

1. Open the file `analytics.js` in the extension directory
2. Find these lines near the top:
   ```javascript
   const GA_MEASUREMENT_ID = 'G-XXXXXXXXXX'; // Replace with actual GA4 Measurement ID
   const GA_API_SECRET = 'XXXXXXXXXXXXXXXXXXXX'; // Replace with actual API Secret
   ```
3. Replace `'G-XXXXXXXXXX'` with your actual Measurement ID from Step 3
4. Replace `'XXXXXXXXXXXXXXXXXXXX'` with your actual API Secret from Step 4
5. Save the file

Example after update:
```javascript
const GA_MEASUREMENT_ID = 'G-ABC1234567';
const GA_API_SECRET = 'xYz123AbC456DeF789GhI012JkL345';
```

## Step 6: Deploy Updated Extension

### For Development/Testing
1. Open your browser and go to `edge://extensions/` or `chrome://extensions/`
2. Enable **Developer mode** (toggle in top right)
3. Click **Load unpacked** and select the extension directory
4. Or, if already loaded, click the **Reload** button for the extension

### For Production Release

**IMPORTANT SECURITY NOTE**: The placeholder credentials in `analytics.js` are not functional and need to be replaced with real credentials before deployment. Consider these approaches:

#### Option A: Direct Replacement (Simple but Less Secure)
1. Replace the placeholder values in `analytics.js` with your actual credentials
2. Commit to a private branch or repository
3. Build and package the extension
4. Upload to extension store

**⚠️ Warning**: This approach exposes credentials in the published extension code. While they're used client-side and rate-limited by Google, it's not ideal for security.

#### Option B: Build-Time Injection (Recommended)
1. Create a `.env` file (add to `.gitignore`) with:
   ```
   GA_MEASUREMENT_ID=G-ABC1234567
   GA_API_SECRET=xYz123AbC456DeF789GhI012JkL345
   ```
2. Use a build script to inject these values during packaging
3. Keep credentials out of version control
4. Only the built package contains real credentials

#### Option C: Configuration File (Alternative)
1. Create an untracked `analytics-config.js` file
2. Import it in `analytics.js`
3. Add `analytics-config.js` to `.gitignore`
4. Document the expected format for other developers

## Step 7: Verify Analytics is Working

### Real-time Verification
1. Install or reload the extension
2. Open the extension popup
3. Click some buttons (e.g., "Check Configuration Assignments")
4. Go back to Google Analytics
5. Navigate to **Reports** → **Realtime**
6. You should see events appearing within 30-60 seconds

### Check Event Names
Look for these events in Realtime reports:
- `page_view` - When popup opens
- `button_click` - When buttons are clicked
- `feature_usage` - When features are used

### Debug Mode (Optional)
1. Open browser DevTools (F12)
2. Go to the **Console** tab
3. Look for `[Analytics]` prefixed messages
4. You'll see confirmation of events being sent

## Step 8: Configure Reports (Optional)

### Create Custom Reports
1. Go to **Reports** → **Engagement** → **Events**
2. Click on event names to see details
3. Create custom reports:
   - Most clicked buttons
   - Feature usage frequency
   - Active users over time

### Set Up Explorations
1. Go to **Explore** tab
2. Create a new exploration
3. Add dimensions: `button_name`, `feature_name`, `extension_version`
4. Add metrics: `Event count`, `Total users`
5. Analyze usage patterns

## Troubleshooting

### Events Not Showing Up

**Check Credentials:**
- Verify Measurement ID format is `G-XXXXXXXXXX`
- Verify API Secret is correct (no extra spaces)
- Ensure both are saved in `analytics.js`

**Check Browser Console:**
- Open DevTools → Console
- Look for `[Analytics]` messages
- Check for any error messages

**Check Network Tab:**
- Open DevTools → Network tab
- Filter by `google-analytics.com`
- Look for POST requests to `mp/collect`
- Check if they're returning 2xx status codes

**Verify Extension Loaded:**
- Ensure extension is properly loaded
- Check for JavaScript errors in console
- Try reloading the extension

### Events Delayed

GA4 Realtime reports can take 30-60 seconds to show events. This is normal.

### No User Interaction Data

Check if analytics is enabled:
1. Open extension popup
2. Click Settings (⋮)
3. Look for "Analytics: Enabled"
4. If it says "Disabled", click to enable it

## Privacy Considerations

### Data Collected
The analytics implementation tracks:
- Anonymous client ID (UUID)
- Button/feature names
- Extension version
- Session information

### Data NOT Collected
NO personal information is collected:
- No device IDs or serial numbers
- No usernames or UPNs
- No Microsoft tokens
- No Intune data (group names, policy names, etc.)
- No browsing history

### User Consent
- Analytics is enabled by default
- Users can opt-out in Settings menu
- Preference is stored locally in browser

### Compliance
This implementation is designed to be:
- GDPR compliant (anonymous data, user control)
- CCPA compliant (no personal data)
- Privacy-focused (minimal data collection)

## Security Best Practices

### Protect Your API Secret
- **DO NOT** commit the API secret to public repositories
- Use environment variables or build-time substitution
- Rotate secrets periodically
- Limit access to GA4 property

### Monitor Access
- Regularly review who has access to your GA4 property
- Remove access for team members who no longer need it
- Use service accounts for automated access

### Review Data
- Regularly audit what data is being collected
- Ensure no PII is accidentally included
- Set up data retention policies in GA4

## Support

If you encounter issues with GA4 setup:

1. **GA4 Documentation**: [Google Analytics Help](https://support.google.com/analytics)
2. **Extension Issues**: Open an issue on [GitHub](../../issues)
3. **Privacy Concerns**: See [ANALYTICS.md](ANALYTICS.md)

## Alternative: Disable Analytics Completely

If you prefer not to use analytics at all:

1. In `analytics.js`, change line:
   ```javascript
   let analyticsEnabled = true; // Default to enabled, user can opt-out
   ```
   to:
   ```javascript
   let analyticsEnabled = false; // Default to disabled
   ```

2. Or simply remove the following from `manifest.json`:
   ```json
   "https://www.google-analytics.com/*"
   ```
   This will prevent any network requests to Google Analytics.

3. Users can also disable it individually in the Settings menu.
