/**
 * UI
 * Custom menu, configuration dialog, and generate action.
 */

/**
 * Add "Mail Merge" menu to the spreadsheet UI.
 * Called automatically when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('Configure…', 'showConfigDialogUI')
    .addSeparator()
    .addItem('Generate', 'runGenerateUI')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Testing')
        .addItem('Create sample data', 'createSampleDataUI')
        .addItem('Run system test', 'runSystemTestUI')
    )
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
 * Full end-to-end sample setup:
 *   1. Create "Sample Data" sheet with 10 dummy records
 *   2. Create Config sheet and default Templates/Output Drive folders
 *   3. Create a sample label template Doc in the Templates folder
 *   4. Save template ID and output folder ID to config
 *   5. Run the mail merge and show the output link
 */
function createSampleDataUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ss.toast('Setting up sample data…', 'Mail Merge', -1);

  // ── Step 1: Sample data sheet ────────────────────────────────────────────
  const sheetName = 'Sample Data';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const headers = ['First Name', 'Last Name', 'Title', 'Organization', 'City'];
  const rows = [
    ['Alice', 'Johnson', 'Director', 'Acme Corp', 'New York'],
    ['Bob', 'Smith', 'Engineer', 'Globex', 'San Francisco'],
    ['Carol', 'Williams', 'Manager', 'Initech', 'Austin'],
    ['David', 'Brown', 'Designer', 'Umbrella', 'Seattle'],
    ['Eva', 'Davis', 'Analyst', 'Stark Industries', 'Chicago'],
    ['Frank', 'Miller', 'VP', 'Wayne Enterprises', 'Boston'],
    ['Grace', 'Wilson', 'Lead', 'Cyberdyne', 'Denver'],
    ['Henry', 'Moore', 'Architect', 'Tyrell Corp', 'Portland'],
    ['Iris', 'Taylor', 'Producer', 'Oscorp', 'Miami'],
    ['Jack', 'Anderson', 'Consultant', 'Soylent Corp', 'Dallas'],
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  for (let i = 0; i < headers.length; i++) sheet.autoResizeColumn(i + 1);
  ss.setActiveSheet(sheet);

  // ── Step 2: Config sheet + default folders ───────────────────────────────
  getConfigSheet();
  const folders = applyDefaultFolders();
  const outputFolderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);

  // ── Step 3: Create sample template Doc ───────────────────────────────────
  const templateDoc = DocumentApp.create('Label Template - Sample');
  const body = templateDoc.getBody();

  // Remove the default empty paragraph, then add a single-cell table
  // whose content serves as the label design.
  body.clear();
  const labelContent = '{{First Name}} {{Last Name}}\n{{Title}}\n{{Organization}}\n{{City}}';
  body.appendTable([[labelContent]]);
  templateDoc.saveAndClose();

  // Move the template into the Templates folder
  const parent = getSpreadsheetParentFolder();
  const templatesFolder = getOrCreateSubfolder(parent, 'Templates');
  DriveApp.getFileById(templateDoc.getId()).moveTo(templatesFolder);

  // ── Step 4: Save config ───────────────────────────────────────────────────
  setConfig(CONFIG_KEYS.TEMPLATE_DOC_ID, templateDoc.getId());

  // ── Step 5: Run the merge ─────────────────────────────────────────────────
  ss.toast('Generating output…', 'Mail Merge', -1);
  const records = getRows(sheet);
  const outputUrl = generateDocument(records, templateDoc.getId(), outputFolderId);

  ss.toast('Done!', 'Mail Merge', 5);
  ui.alert(
    'Sample setup complete',
    'All steps completed successfully:\n\n' +
    '✓ "Sample Data" sheet created (10 records)\n' +
    '✓ Config sheet created\n' +
    '✓ Templates folder: ' + folders.templatesUrl + '\n' +
    '✓ Output folder: ' + folders.outputUrl + '\n' +
    '✓ Sample template Doc created\n' +
    '✓ Mail merge output generated\n\n' +
    'Output file:\n' + outputUrl,
    ui.ButtonSet.OK
  );
}

/**
 * Run a system test: checks config, folder access, and template access.
 * Shows a summary dialog with pass/fail results.
 */
function runSystemTestUI() {
  const ui = SpreadsheetApp.getUi();
  const results = [];

  // Test 1: Config sheet exists
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  results.push({ name: 'Config sheet exists', pass: !!configSheet });

  // Test 2: Output folder configured
  const outputFolderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);
  results.push({ name: 'Output folder configured', pass: !!outputFolderId });

  // Test 3: Output folder accessible
  if (outputFolderId) {
    try {
      DriveApp.getFolderById(outputFolderId);
      results.push({ name: 'Output folder accessible', pass: true });
    } catch (e) {
      results.push({ name: 'Output folder accessible', pass: false, detail: e.message });
    }
  } else {
    results.push({ name: 'Output folder accessible', pass: false, detail: 'Not configured' });
  }

  // Test 4: At least one template configured
  const docId = getConfig(CONFIG_KEYS.TEMPLATE_DOC_ID);
  const slidesId = getConfig(CONFIG_KEYS.TEMPLATE_SLIDES_ID);
  results.push({ name: 'Template configured', pass: !!(docId || slidesId) });

  // Test 5: Template accessible (whichever is configured)
  const templateId = docId || slidesId;
  if (templateId) {
    try {
      DriveApp.getFileById(templateId);
      results.push({ name: 'Template file accessible', pass: true });
    } catch (e) {
      results.push({ name: 'Template file accessible', pass: false, detail: e.message });
    }
  } else {
    results.push({ name: 'Template file accessible', pass: false, detail: 'Not configured' });
  }

  // Test 6: Active sheet has data rows
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const hasData = sheet.getLastRow() >= 2 && sheet.getLastColumn() >= 1;
  results.push({ name: 'Active sheet has data', pass: hasData });

  // Build summary
  const passed = results.filter(r => r.pass).length;
  const lines = results.map(r => (r.pass ? '✓' : '✗') + ' ' + r.name + (r.detail ? ' (' + r.detail + ')' : ''));
  const summary = passed + '/' + results.length + ' checks passed\n\n' + lines.join('\n');

  ui.alert('System Test Results', summary, ui.ButtonSet.OK);
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
