/**
 * Sheet Reader
 * Reads headers and data rows from the active Google Sheet.
 */

/**
 * Get column headers from the first row of a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {string[]} Array of header strings
 */
function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headerRow.map(h => h.toString().trim());
}

/**
 * Get data rows from a sheet as an array of record objects.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} [options]
 * @param {number} [options.limit] - Maximum number of rows to return
 * @param {string} [options.statusColumn] - Header name of the column to filter on
 * @param {string} [options.statusValue] - Only include rows where statusColumn equals this value
 * @returns {Object[]} Array of records mapping header → cell value
 */
function getRows(sheet, options) {
  options = options || {};

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol === 0) return [];

  const headers = getHeaders(sheet);
  const dataValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const records = [];

  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    const record = {};

    for (let j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j] !== undefined ? row[j] : '';
    }

    // Apply status filter
    if (options.statusColumn && options.statusValue !== undefined) {
      const cellValue = record[options.statusColumn];
      if (cellValue === undefined || cellValue.toString() !== options.statusValue.toString()) {
        continue;
      }
    }

    records.push(record);

    if (options.limit && records.length >= options.limit) break;
  }

  return records;
}
