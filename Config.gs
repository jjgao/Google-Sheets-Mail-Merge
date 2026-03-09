/**
 * Configuration Management
 * Handles getting and setting configuration values using a Config sheet tab.
 */

const CONFIG_KEYS = {
  TEMPLATE_DOC_ID: 'TEMPLATE_DOC_ID',
  TEMPLATE_SLIDES_ID: 'TEMPLATE_SLIDES_ID',
  OUTPUT_FOLDER_ID: 'OUTPUT_FOLDER_ID',
};

const CONFIG_SHEET_NAME = 'Config';

const CONFIG_LABELS = {
  TEMPLATE_DOC_ID: 'Template Doc ID',
  TEMPLATE_SLIDES_ID: 'Template Slides ID',
  OUTPUT_FOLDER_ID: 'Output Folder ID',
};

const CONFIG_DESCRIPTIONS = {
  TEMPLATE_DOC_ID: 'Google Doc ID for label/badge template (from URL) — optional if Slides template is set',
  TEMPLATE_SLIDES_ID: 'Google Slides ID for label/badge template (from URL) — optional if Doc template is set',
  OUTPUT_FOLDER_ID: 'Google Drive Folder ID where output files will be saved — required',
};

/**
 * Get a configuration value from the Config sheet.
 * @param {string} key - Configuration key (use CONFIG_KEYS constants)
 * @returns {string|null} Configuration value or null if not set
 */
function getConfig(key) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

    if (!sheet) {
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const label = CONFIG_LABELS[key];

    // Row 0: header; Row 1+: data rows
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === label) {
        const value = data[i][1];
        if (value && value.toString().trim() !== '') {
          return value.toString().trim();
        }
        break;
      }
    }

    return null;

  } catch (error) {
    Logger.log('Error reading config for ' + key + ': ' + error.message);
    return null;
  }
}

/**
 * Set a configuration value in the Config sheet.
 * @param {string} key - Configuration key
 * @param {string} value - Configuration value
 */
function setConfig(key, value) {
  try {
    const sheet = getConfigSheet();
    const data = sheet.getDataRange().getValues();
    const label = CONFIG_LABELS[key];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === label) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
  } catch (error) {
    throw new Error('Failed to set config ' + key + ': ' + error.message);
  }
}

/**
 * Set multiple configuration values at once.
 * @param {Object} configObject - Object mapping CONFIG_KEYS to values
 */
function setMultipleConfig(configObject) {
  for (const key in configObject) {
    setConfig(key, configObject[key]);
  }
}

/**
 * Get all configuration values.
 * @returns {Object} All configuration values
 */
function getAllConfig() {
  const config = {};
  for (const key in CONFIG_KEYS) {
    config[key] = getConfig(CONFIG_KEYS[key]);
  }
  return config;
}

/**
 * Validate that required configuration is set.
 * Requires OUTPUT_FOLDER_ID and at least one template ID.
 * @returns {{ isValid: boolean, missing: string[] }}
 */
function validateConfig() {
  const missing = [];

  if (!getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID)) {
    missing.push(CONFIG_KEYS.OUTPUT_FOLDER_ID);
  }

  const hasDoc = !!getConfig(CONFIG_KEYS.TEMPLATE_DOC_ID);
  const hasSlides = !!getConfig(CONFIG_KEYS.TEMPLATE_SLIDES_ID);

  if (!hasDoc && !hasSlides) {
    missing.push(CONFIG_KEYS.TEMPLATE_DOC_ID);
    missing.push(CONFIG_KEYS.TEMPLATE_SLIDES_ID);
  }

  return {
    isValid: missing.length === 0,
    missing: missing,
  };
}

/**
 * Get the parent Drive folder of the active spreadsheet.
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getSpreadsheetParentFolder() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  return file.getParents().next();
}

/**
 * Get a named subfolder within a parent folder, creating it if it doesn't exist.
 * @param {GoogleAppsScript.Drive.Folder} parentFolder
 * @param {string} name
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateSubfolder(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(name);
}

/**
 * Create default "Templates" and "Output" folders under the spreadsheet's
 * parent folder (if they don't already exist), and save the Output folder ID
 * to config.
 * @returns {{ templatesUrl: string, outputUrl: string }}
 */
function applyDefaultFolders() {
  const parent = getSpreadsheetParentFolder();
  const templatesFolder = getOrCreateSubfolder(parent, 'Templates');
  const outputFolder = getOrCreateSubfolder(parent, 'Output');

  // Only set OUTPUT_FOLDER_ID if it's not already configured
  if (!getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID)) {
    setConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID, outputFolder.getId());
  }

  return {
    templatesUrl: templatesFolder.getUrl(),
    outputUrl: outputFolder.getUrl(),
  };
}

/**
 * Get or create the Config sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getConfigSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
    initializeConfigSheet(sheet);
  }

  return sheet;
}

/**
 * Initialize the Config sheet with headers and setting rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function initializeConfigSheet(sheet) {
  const headers = ['Setting', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 400);

  const data = [];
  for (const key in CONFIG_KEYS) {
    const configKey = CONFIG_KEYS[key];
    data.push([
      CONFIG_LABELS[configKey],
      '',
      CONFIG_DESCRIPTIONS[configKey] || '',
    ]);
  }

  sheet.getRange(2, 1, data.length, 3).setValues(data);
}
