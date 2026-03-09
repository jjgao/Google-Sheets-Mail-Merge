/**
 * Google Sheets Mail Merge
 * Entry point and application info.
 */

const APP_INFO = {
  name: 'Google Sheets Mail Merge',
  version: '1.0.0',
  description: 'Merge Google Sheet data into tiled label/badge layouts saved to Google Drive.',
};

/**
 * Returns the current app version string.
 * @returns {string}
 */
function getVersion() {
  return APP_INFO.version;
}

/**
 * Initialize the application: creates the Config sheet if it does not exist.
 * Run this once when setting up the add-on in a new spreadsheet.
 */
function initializeApp() {
  getConfigSheet(); // creates Config sheet if missing
  SpreadsheetApp.getActiveSpreadsheet().toast(
    APP_INFO.name + ' v' + APP_INFO.version + ' initialized.',
    'Mail Merge',
    5
  );
}
