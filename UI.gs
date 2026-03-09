/**
 * UI
 * Custom menu, configuration dialog, and generate action.
 */

/**
 * Add "Mail Merge" menu to the spreadsheet UI.
 * Called automatically when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail Merge')
    .addItem('Configure…', 'showConfigDialogUI')
    .addSeparator()
    .addItem('Generate', 'runGenerateUI')
    .addToUi();
}

/**
 * Show the configuration dialog.
 */
function showConfigDialogUI() {
  const config = getAllConfig();
  const html = HtmlService.createHtmlOutput(buildConfigHtml(config))
    .setWidth(480)
    .setHeight(300)
    .setTitle('Mail Merge Configuration');
  SpreadsheetApp.getUi().showModalDialog(html, 'Mail Merge Configuration');
}

/**
 * Build the HTML string for the config dialog.
 * @param {Object} config
 * @returns {string}
 */
function buildConfigHtml(config) {
  return '<!DOCTYPE html><html><body>' +
    '<style>body{font-family:sans-serif;font-size:13px;padding:16px}' +
    'label{display:block;margin-top:12px;font-weight:bold}' +
    'input{width:100%;box-sizing:border-box;padding:6px;margin-top:4px;border:1px solid #ccc;border-radius:3px}' +
    '.hint{font-size:11px;color:#666;margin-top:2px}' +
    'button{margin-top:16px;padding:8px 20px;background:#4285f4;color:#fff;border:none;border-radius:3px;cursor:pointer}' +
    '</style>' +
    '<form>' +
    '<label>Template Doc ID (optional)</label>' +
    '<input name="TEMPLATE_DOC_ID" value="' + (config.TEMPLATE_DOC_ID || '') + '" />' +
    '<div class="hint">Google Doc ID from the URL — leave blank if using Slides template</div>' +
    '<label>Template Slides ID (optional)</label>' +
    '<input name="TEMPLATE_SLIDES_ID" value="' + (config.TEMPLATE_SLIDES_ID || '') + '" />' +
    '<div class="hint">Google Slides ID from the URL — leave blank if using Doc template</div>' +
    '<label>Output Folder ID (required)</label>' +
    '<input name="OUTPUT_FOLDER_ID" value="' + (config.OUTPUT_FOLDER_ID || '') + '" />' +
    '<div class="hint">Google Drive folder ID where output files will be saved</div>' +
    '<button type="button" onclick="saveConfig()">Save</button>' +
    '</form>' +
    '<script>' +
    'function saveConfig(){' +
    'var form=document.querySelector("form");' +
    'var data={};' +
    'for(var i=0;i<form.elements.length;i++){var el=form.elements[i];if(el.name)data[el.name]=el.value;}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).saveConfigFromDialog(data);' +
    '}' +
    '</script>' +
    '</body></html>';
}

/**
 * Save configuration values submitted from the config dialog.
 * @param {Object} data - Object with CONFIG_KEYS as keys
 */
function saveConfigFromDialog(data) {
  for (const key in data) {
    if (CONFIG_KEYS[key] !== undefined) {
      setConfig(key, data[key]);
    }
  }
}

/**
 * Run the mail merge generate action.
 * Reads config and sheet data, calls the appropriate generator, shows result.
 */
function runGenerateUI() {
  const ui = SpreadsheetApp.getUi();

  // Validate config
  const validation = validateConfig();
  if (!validation.isValid) {
    ui.alert('Configuration incomplete. Missing: ' + validation.missing.join(', ') + '. Please configure via Mail Merge → Configure…');
    return;
  }

  const config = getAllConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const records = getRows(sheet);

  if (records.length === 0) {
    ui.alert('No data rows found in the active sheet.');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Generating output…', 'Mail Merge', -1);

  try {
    let outputUrl;

    if (config.TEMPLATE_DOC_ID) {
      outputUrl = generateDocument(records, config.TEMPLATE_DOC_ID, config.OUTPUT_FOLDER_ID);
    } else {
      outputUrl = generateSlides(records, config.TEMPLATE_SLIDES_ID, config.OUTPUT_FOLDER_ID);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Done!', 'Mail Merge', 5);
    ui.alert('Output saved to Drive:\n' + outputUrl);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error', 'Mail Merge', 5);
    ui.alert('Error generating output: ' + error.message);
  }
}
