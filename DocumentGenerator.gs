/**
 * Document Generator
 * Clones a Google Doc template, tiles records into a grid layout,
 * substitutes {{placeholder}} fields, and saves the output to Drive.
 *
 * Template convention:
 *   The template Doc must contain a table. The table dimensions define the
 *   grid: a 2×5 table means 2 columns × 5 rows = 10 labels per page. The
 *   content of cell [0][0] is used as the label text template for all cells.
 *
 *   If no table is found in the template, the body text is used as the
 *   label template with a 1×1 grid (one record per page).
 */

/**
 * Detect grid dimensions from the first table in a Google Document.
 * The table's own dimensions (rows × cols) define the grid per page.
 * @param {GoogleAppsScript.Document.Document} doc
 * @returns {{ cols: number, rows: number }}
 */
function detectDocGrid(doc) {
  const body = doc.getBody();

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      const rows = table.getNumRows();
      const cols = rows > 0 ? table.getRow(0).getNumCells() : 1;
      return {
        cols: Math.max(1, cols),
        rows: Math.max(1, rows),
      };
    }
  }

  return { cols: 1, rows: 1 };
}

/**
 * Generate a tiled Google Doc from a template and a list of records.
 *
 * Steps:
 *   1. Copy the template Doc to the output folder.
 *   2. Extract the template text from the first table cell (or body text).
 *   3. Clear the document body.
 *   4. For each page-worth of records, append a table via appendTable(cells)
 *      with all placeholder values substituted, then a page break.
 *
 * @param {Object[]} records - Array of record objects from SheetReader
 * @param {string} templateDocId - Google Doc ID of the template
 * @param {string} outputFolderId - Drive folder ID for output
 * @returns {string} URL of the generated output file
 */
function generateDocument(records, templateDocId, outputFolderId) {
  const outputName = 'Labels - ' + Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  const outputFile = DriveApp.getFileById(templateDocId)
    .makeCopy(outputName, DriveApp.getFolderById(outputFolderId));
  const doc = DocumentApp.openById(outputFile.getId());
  const body = doc.getBody();

  // Extract grid and template text before modifying the document
  const grid = detectDocGrid(doc);
  let templateText = '';
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      templateText = child.asTable().getCell(0, 0).getText();
      break;
    }
  }
  if (!templateText) {
    templateText = body.getText();
  }

  // Clear body and rebuild with tiled records
  body.clear();

  let recordIndex = 0;
  let firstPage = true;

  while (recordIndex < records.length) {
    if (!firstPage) {
      body.appendPageBreak();
    }
    firstPage = false;

    // Build cols × rows 2D cell array for this page
    const cells = [];
    for (let row = 0; row < grid.rows; row++) {
      const rowData = [];
      for (let col = 0; col < grid.cols; col++) {
        rowData.push(recordIndex < records.length
          ? substitutePlaceholders(templateText, records[recordIndex++])
          : '');
      }
      cells.push(rowData);
    }

    body.appendTable(cells);
  }

  doc.saveAndClose();
  return outputFile.getUrl();
}
