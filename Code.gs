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
 * Initialize the application:
 * - Creates the Config sheet if it does not exist
 * - Creates default Templates and Output folders under the spreadsheet's parent folder
 * - Saves the Output folder ID to config
 *
 * Run this once when setting up the add-on in a new spreadsheet.
 */
function initializeApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getConfigSheet(); // creates Config sheet if missing

  const folders = applyDefaultFolders();

  ss.toast(
    APP_INFO.name + ' v' + APP_INFO.version + ' initialized. ' +
    'Output folder ready. Use Mail Merge → Configure… to set your template.',
    'Mail Merge',
    8
  );

  SpreadsheetApp.getUi().alert(
    'Mail Merge Initialized',
    'Default folders created (or already existed):\n\n' +
    'Templates: ' + folders.templatesUrl + '\n\n' +
    'Output:    ' + folders.outputUrl + '\n\n' +
    'Place your template Doc or Slides file in the Templates folder, ' +
    'then set its ID via Mail Merge → Configure…',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
