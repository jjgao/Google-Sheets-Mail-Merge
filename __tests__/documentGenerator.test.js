describe('detectDocGrid', () => {
  function makeDoc({ numTableRows = 5, numTableCols = 2, hasTable = true } = {}) {
    const mockTable = {
      getType: jest.fn(() => DocumentApp.ElementType.TABLE),
      asTable: function() { return this; },
      getNumRows: jest.fn(() => numTableRows),
      getRow: jest.fn(() => ({
        getNumCells: jest.fn(() => numTableCols),
      })),
      getCell: jest.fn(() => ({
        getText: jest.fn(() => '{{Name}}'),
        setText: jest.fn(),
        getNumChildren: jest.fn(() => 0),
      })),
    };

    const children = hasTable ? [mockTable] : [];

    const mockBody = {
      getNumChildren: jest.fn(() => children.length),
      getChild: jest.fn((i) => children[i]),
      getPageWidth: jest.fn(() => 612),
      getPageHeight: jest.fn(() => 792),
      getMarginTop: jest.fn(() => 36),
      getMarginBottom: jest.fn(() => 36),
      getMarginLeft: jest.fn(() => 36),
      getMarginRight: jest.fn(() => 36),
      getText: jest.fn(() => ''),
      clear: jest.fn(),
      appendTable: jest.fn(),
      appendPageBreak: jest.fn(),
    };

    return { getBody: jest.fn(() => mockBody) };
  }

  test('returns 1×1 when no table is found', () => {
    const doc = makeDoc({ hasTable: false });
    expect(detectDocGrid(doc)).toEqual({ cols: 1, rows: 1 });
  });

  test('returns table dimensions as grid', () => {
    const doc = makeDoc({ numTableRows: 5, numTableCols: 2 });
    expect(detectDocGrid(doc)).toEqual({ cols: 2, rows: 5 });
  });

  test('returns 1×1 for a 1-cell table', () => {
    const doc = makeDoc({ numTableRows: 1, numTableCols: 1 });
    expect(detectDocGrid(doc)).toEqual({ cols: 1, rows: 1 });
  });

  test('returns at least 1 col and 1 row', () => {
    const doc = makeDoc({ numTableRows: 0, numTableCols: 0 });
    const { cols, rows } = detectDocGrid(doc);
    expect(cols).toBeGreaterThanOrEqual(1);
    expect(rows).toBeGreaterThanOrEqual(1);
  });
});

describe('generateDocument', () => {
  function makeTemplateCellMock(text = '{{Name}}') {
    return {
      getText: jest.fn(() => text),
      setText: jest.fn(),
      getNumChildren: jest.fn(() => 0),
    };
  }

  function makeDocMock(templateText = '{{Name}}') {
    const templateCell = makeTemplateCellMock(templateText);
    const templateRow = {
      getNumCells: jest.fn(() => 1),
      getCell: jest.fn(() => templateCell),
    };
    const mockTable = {
      getType: jest.fn(() => DocumentApp.ElementType.TABLE),
      asTable: function() { return this; },
      getNumRows: jest.fn(() => 1),
      getRow: jest.fn(() => templateRow),
      getCell: jest.fn(() => templateCell),
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
      getText: jest.fn(() => templateText),
      clear: jest.fn(),
      appendTable: jest.fn(),
      appendPageBreak: jest.fn(),
    };

    return {
      getBody: jest.fn(() => mockBody),
      saveAndClose: jest.fn(),
    };
  }

  function makeOutputFileMock() {
    return {
      getId: jest.fn(() => 'output-doc-id'),
      getUrl: jest.fn(() => 'https://docs.google.com/output-doc-id'),
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
    const url = generateDocument([{ Name: 'Alice' }], 'template-id', 'folder-id');
    expect(typeof url).toBe('string');
    expect(url).toContain('http');
  });

  test('calls makeCopy on the template file', () => {
    generateDocument([{ Name: 'Alice' }], 'template-id', 'folder-id');
    expect(DriveApp.getFileById).toHaveBeenCalledWith('template-id');
    expect(DriveApp.getFileById.mock.results[0].value.makeCopy).toHaveBeenCalled();
  });

  test('output file name includes "Labels"', () => {
    generateDocument([{ Name: 'Alice' }], 'template-id', 'folder-id');
    const copyName = DriveApp.getFileById.mock.results[0].value.makeCopy.mock.calls[0][0];
    expect(copyName).toMatch(/Labels/);
  });

  test('opens the copied document', () => {
    generateDocument([{ Name: 'Alice' }], 'template-id', 'folder-id');
    expect(DocumentApp.openById).toHaveBeenCalledWith('output-doc-id');
  });

  test('clears the body before rebuilding', () => {
    generateDocument([{ Name: 'Alice' }], 'template-id', 'folder-id');
    const body = DocumentApp.openById.mock.results[0].value.getBody();
    expect(body.clear).toHaveBeenCalled();
  });

  test('appends one table per page of records', () => {
    // 1×1 grid, 3 records → 3 pages → 3 appendTable calls
    generateDocument(
      [{ Name: 'Alice' }, { Name: 'Bob' }, { Name: 'Carol' }],
      'template-id', 'folder-id'
    );
    const body = DocumentApp.openById.mock.results[0].value.getBody();
    expect(body.appendTable).toHaveBeenCalledTimes(3);
  });

  test('appends page breaks between pages', () => {
    // 3 records → 2 page breaks (between page 1→2 and 2→3)
    generateDocument(
      [{ Name: 'Alice' }, { Name: 'Bob' }, { Name: 'Carol' }],
      'template-id', 'folder-id'
    );
    const body = DocumentApp.openById.mock.results[0].value.getBody();
    expect(body.appendPageBreak).toHaveBeenCalledTimes(2);
  });

  test('saves and closes the document', () => {
    generateDocument([{ Name: 'Alice' }], 'template-id', 'folder-id');
    expect(DocumentApp.openById.mock.results[0].value.saveAndClose).toHaveBeenCalled();
  });

  test('substitutes placeholders into cell content', () => {
    generateDocument([{ Name: 'Alice' }], 'template-id', 'folder-id');
    const body = DocumentApp.openById.mock.results[0].value.getBody();
    const cells = body.appendTable.mock.calls[0][0];
    expect(cells[0][0]).toBe('Alice');
  });
});
