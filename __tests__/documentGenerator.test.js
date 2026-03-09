describe('detectDocGrid', () => {
  function makeDoc({ pageWidth = 612, pageHeight = 792, marginTop = 36, marginBottom = 36, marginLeft = 36, marginRight = 36, tableWidth = 540, numTableCols = 3, numTableRows = 5, hasTable = true } = {}) {
    const usableWidth = pageWidth - marginLeft - marginRight;
    const usableHeight = pageHeight - marginTop - marginBottom;

    const mockTable = {
      getType: () => DocumentApp.ElementType.TABLE,
      asTable: function() { return this; },
      getNumRows: jest.fn(() => numTableRows),
      getWidth: jest.fn(() => tableWidth),
      getRow: jest.fn(() => ({
        getNumCells: jest.fn(() => numTableCols),
      })),
    };

    const children = hasTable ? [mockTable] : [];

    const mockBody = {
      getPageWidth: jest.fn(() => pageWidth),
      getPageHeight: jest.fn(() => pageHeight),
      getMarginTop: jest.fn(() => marginTop),
      getMarginBottom: jest.fn(() => marginBottom),
      getMarginLeft: jest.fn(() => marginLeft),
      getMarginRight: jest.fn(() => marginRight),
      getNumChildren: jest.fn(() => children.length),
      getChild: jest.fn((i) => children[i]),
    };

    return { getBody: jest.fn(() => mockBody) };
  }

  test('returns 1×1 when no table is found', () => {
    const doc = makeDoc({ hasTable: false });
    expect(detectDocGrid(doc)).toEqual({ cols: 1, rows: 1 });
  });

  test('computes cols from table and page width', () => {
    // tableWidth=540, numTableCols=3 → cellWidth=180; usableWidth=540 → cols=3
    const doc = makeDoc({ tableWidth: 540, numTableCols: 3 });
    const { cols } = detectDocGrid(doc);
    expect(cols).toBe(3);
  });

  test('returns at least 1 col and 1 row', () => {
    const doc = makeDoc({ tableWidth: 10000, numTableCols: 1, numTableRows: 1 });
    const { cols, rows } = detectDocGrid(doc);
    expect(cols).toBeGreaterThanOrEqual(1);
    expect(rows).toBeGreaterThanOrEqual(1);
  });
});

describe('generateDocument', () => {
  function makeOutputFileMock() {
    return {
      getId: jest.fn(() => 'output-doc-id'),
      getUrl: jest.fn(() => 'https://docs.google.com/output-doc-id'),
    };
  }

  function makeDocMock(templateText = '{{Name}}') {
    const templateCell = {
      getText: jest.fn(() => templateText),
      setText: jest.fn(),
      getNumChildren: jest.fn(() => 0),
    };
    const templateRow = {
      getNumCells: jest.fn(() => 1),
      getCell: jest.fn(() => templateCell),
      appendTableCell: jest.fn(() => ({
        setText: jest.fn(),
        getText: jest.fn(() => ''),
        getNumChildren: jest.fn(() => 0),
      })),
    };
    const mockTable = {
      getType: jest.fn(() => DocumentApp.ElementType.TABLE),
      asTable: function() { return this; },
      getNumRows: jest.fn(() => 1),
      getRow: jest.fn(() => templateRow),
      getWidth: jest.fn(() => 540),
      appendTableRow: jest.fn(() => ({
        getNumCells: jest.fn(() => 0),
        appendTableCell: jest.fn(() => ({
          setText: jest.fn(),
          getText: jest.fn(() => ''),
        })),
      })),
      removeRow: jest.fn(),
    };

    const mockBody = {
      getPageWidth: jest.fn(() => 612),
      getPageHeight: jest.fn(() => 792),
      getMarginTop: jest.fn(() => 36),
      getMarginBottom: jest.fn(() => 36),
      getMarginLeft: jest.fn(() => 36),
      getMarginRight: jest.fn(() => 36),
      getNumChildren: jest.fn(() => 1),
      getChild: jest.fn(() => mockTable),
      appendPageBreak: jest.fn(),
      appendTable: jest.fn(() => mockTable),
      getText: jest.fn(() => templateText),
      setText: jest.fn(),
    };

    return {
      getBody: jest.fn(() => mockBody),
      saveAndClose: jest.fn(),
    };
  }

  beforeEach(() => {
    const outputFile = makeOutputFileMock();
    DriveApp.getFolderById.mockReturnValue({ getId: jest.fn(() => 'folder-id') });
    DriveApp.getFileById.mockReturnValue({
      makeCopy: jest.fn(() => outputFile),
    });
    DocumentApp.openById.mockReturnValue(makeDocMock());
  });

  test('returns a URL string', () => {
    const records = [{ Name: 'Alice' }];
    const url = generateDocument(records, 'template-id', 'folder-id');
    expect(typeof url).toBe('string');
    expect(url).toContain('http');
  });

  test('calls makeCopy on the template file', () => {
    const records = [{ Name: 'Alice' }];
    generateDocument(records, 'template-id', 'folder-id');
    expect(DriveApp.getFileById).toHaveBeenCalledWith('template-id');
    const templateFileMock = DriveApp.getFileById.mock.results[0].value;
    expect(templateFileMock.makeCopy).toHaveBeenCalled();
  });

  test('output file name includes "Labels"', () => {
    const records = [{ Name: 'Alice' }];
    generateDocument(records, 'template-id', 'folder-id');
    const templateFileMock = DriveApp.getFileById.mock.results[0].value;
    const copyName = templateFileMock.makeCopy.mock.calls[0][0];
    expect(copyName).toMatch(/Labels/);
  });

  test('opens the copied document', () => {
    const records = [{ Name: 'Alice' }];
    generateDocument(records, 'template-id', 'folder-id');
    expect(DocumentApp.openById).toHaveBeenCalledWith('output-doc-id');
  });

  test('saves and closes the document', () => {
    const records = [{ Name: 'Alice' }];
    generateDocument(records, 'template-id', 'folder-id');
    const doc = DocumentApp.openById.mock.results[0].value;
    expect(doc.saveAndClose).toHaveBeenCalled();
  });
});
