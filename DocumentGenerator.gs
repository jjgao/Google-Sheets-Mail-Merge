/**
 * Document Generator
 * Clones a Google Doc template, tiles records into a grid layout,
 * substitutes {{placeholder}} fields, and saves the output to Drive.
 *
 * Template convention:
 *   The template Doc must contain exactly one table. That table defines
 *   one label cell (the content of cell [0][0]). The generator detects
 *   how many columns fit per page from the table's column widths and the
 *   page width, and how many rows fit per page from the row heights and
 *   page height. It then fills in the grid row-by-row, page-by-page.
 *
 *   If no table is found, the entire body content is treated as a single
 *   label cell and a 1×1 grid (one record per page) is used.
 */

/**
 * Detect the grid dimensions from a Google Document.
 * Reads the first table in the doc to determine cols×rows per page.
 * Falls back to 1×1 if no table is found.
 * @param {GoogleAppsScript.Document.Document} doc
 * @returns {{ cols: number, rows: number }}
 */
function detectDocGrid(doc) {
  const body = doc.getBody();
  const pageWidth = body.getPageWidth();
  const pageHeight = body.getPageHeight();
  const marginTop = body.getMarginTop();
  const marginBottom = body.getMarginBottom();
  const marginLeft = body.getMarginLeft();
  const marginRight = body.getMarginRight();

  const usableWidth = pageWidth - marginLeft - marginRight;
  const usableHeight = pageHeight - marginTop - marginBottom;

  // Search for the first table element in the body
  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      const numTableRows = table.getNumRows();
      if (numTableRows === 0) break;

      const firstRow = table.getRow(0);
      const numTableCols = firstRow.getNumCells();
      if (numTableCols === 0) break;

      // Estimate cell dimensions from the table dimensions
      // getWidth() returns the width of the table; distribute evenly
      const tableWidth = table.getWidth ? table.getWidth() : usableWidth;
      const cellWidth = tableWidth / numTableCols;

      // Estimate cell height from row height (not always available)
      // Use table height divided by number of rows as approximation
      const cellHeight = usableHeight / numTableRows;

      const cols = Math.max(1, Math.floor(usableWidth / cellWidth));
      const rows = Math.max(1, Math.floor(usableHeight / cellHeight));

      return { cols, rows };
    }
  }

  // No table found — treat as single-label-per-page
  return { cols: 1, rows: 1 };
}

/**
 * Substitute all {{placeholder}} tokens in all text runs within a table cell.
 * @param {GoogleAppsScript.Document.TableCell} cell
 * @param {Object} record
 */
function substituteInDocCell(cell, record) {
  const numParagraphs = cell.getNumChildren();
  for (let p = 0; p < numParagraphs; p++) {
    const para = cell.getChild(p);
    const text = para.asText ? para.asText() : null;
    if (!text) continue;
    const raw = text.getText();
    const replaced = substitutePlaceholders(raw, record);
    if (replaced !== raw) {
      text.setText(replaced);
    }
  }
}

/**
 * Copy the content of one table cell into another, preserving text.
 * @param {GoogleAppsScript.Document.TableCell} sourceCell
 * @param {GoogleAppsScript.Document.TableCell} targetCell
 */
function copyDocCell(sourceCell, targetCell) {
  // Clear target
  while (targetCell.getNumChildren() > 0) {
    targetCell.removeChild(targetCell.getChild(0));
  }

  // Copy paragraphs from source
  const numChildren = sourceCell.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = sourceCell.getChild(i);
    targetCell.appendParagraph(child.asText ? child.asText().getText() : '');
  }
}

/**
 * Generate a tiled Google Doc from a template and a list of records.
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

  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const templateFile = DriveApp.getFileById(templateDocId);
  const outputFile = templateFile.makeCopy(outputName, outputFolder);
  const outputFileId = outputFile.getId();

  const doc = DocumentApp.openById(outputFileId);
  const body = doc.getBody();
  const grid = detectDocGrid(doc);
  const cellsPerPage = grid.cols * grid.rows;

  // Find the template table (first table in the doc)
  let templateTable = null;
  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      templateTable = child.asTable();
      break;
    }
  }

  if (!templateTable) {
    // No table: do simple full-body substitution (1 record per page)
    for (let r = 0; r < records.length; r++) {
      if (r > 0) {
        body.appendPageBreak();
        // Duplicate body content for each record (simplified: just substitute in place)
      }
      const rawText = body.getText();
      const substituted = substitutePlaceholders(rawText, records[r]);
      body.setText(substituted);
    }
    doc.saveAndClose();
    return outputFile.getUrl();
  }

  // Get the template cell content (row 0, col 0 of the template table)
  const templateCell = templateTable.getRow(0).getCell(0);
  const templateText = templateCell.getText();

  // Clear the table body (we will repopulate it)
  // Remove all rows except first (which we keep as a reference)
  while (templateTable.getNumRows() > 1) {
    templateTable.removeRow(templateTable.getNumRows() - 1);
  }
  while (templateTable.getRow(0).getNumCells() > 1) {
    templateTable.getRow(0).removeCell(templateTable.getRow(0).getNumCells() - 1);
  }

  let recordIndex = 0;
  let isFirstPage = true;

  while (recordIndex < records.length) {
    if (!isFirstPage) {
      body.appendPageBreak();
      // appendTable not available directly after page break in all GAS versions;
      // we replicate the table by inserting a new table element
    }

    // For the first page, reuse the existing templateTable.
    // For subsequent pages, insert a new table.
    let currentTable;
    if (isFirstPage) {
      currentTable = templateTable;
      isFirstPage = false;
    } else {
      currentTable = body.appendTable();
    }

    // Build grid.rows × grid.cols cells, filling with records
    for (let row = 0; row < grid.rows; row++) {
      let tableRow;
      if (row < currentTable.getNumRows()) {
        tableRow = currentTable.getRow(row);
      } else {
        tableRow = currentTable.appendTableRow();
      }

      for (let col = 0; col < grid.cols; col++) {
        let cell;
        if (col < tableRow.getNumCells()) {
          cell = tableRow.getCell(col);
        } else {
          cell = tableRow.appendTableCell();
        }

        if (recordIndex < records.length) {
          const substituted = substitutePlaceholders(templateText, records[recordIndex]);
          cell.setText(substituted);
          recordIndex++;
        } else {
          // Remainder cell — leave blank
          cell.setText('');
        }
      }
    }
  }

  doc.saveAndClose();
  return outputFile.getUrl();
}
