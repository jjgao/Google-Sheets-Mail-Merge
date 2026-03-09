function makeSheet(headers, rows) {
  const allRows = [headers, ...rows];
  return {
    getLastColumn: jest.fn(() => headers.length),
    getLastRow: jest.fn(() => allRows.length),
    getRange: jest.fn((startRow, startCol, numRows, numCols) => ({
      getValues: jest.fn(() => {
        return allRows.slice(startRow - 1, startRow - 1 + numRows).map(r => r.slice(startCol - 1, startCol - 1 + numCols));
      }),
    })),
  };
}

describe('getHeaders', () => {
  test('returns headers from first row', () => {
    const sheet = makeSheet(['First Name', 'Last Name', 'City'], []);
    expect(getHeaders(sheet)).toEqual(['First Name', 'Last Name', 'City']);
  });

  test('trims whitespace from headers', () => {
    const sheet = makeSheet([' Name ', ' Email '], []);
    expect(getHeaders(sheet)).toEqual(['Name', 'Email']);
  });

  test('returns empty array for empty sheet', () => {
    const sheet = {
      getLastColumn: jest.fn(() => 0),
      getRange: jest.fn(),
    };
    expect(getHeaders(sheet)).toEqual([]);
  });
});

describe('getRows', () => {
  test('returns records as objects keyed by header', () => {
    const sheet = makeSheet(
      ['First Name', 'City'],
      [['Alice', 'NYC'], ['Bob', 'LA']],
    );
    const rows = getRows(sheet);
    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual({ 'First Name': 'Alice', City: 'NYC' });
    expect(rows[1]).toEqual({ 'First Name': 'Bob', City: 'LA' });
  });

  test('returns empty array when no data rows', () => {
    const sheet = makeSheet(['Name'], []);
    expect(getRows(sheet)).toEqual([]);
  });

  test('respects limit option', () => {
    const sheet = makeSheet(
      ['Name'],
      [['Alice'], ['Bob'], ['Carol']],
    );
    const rows = getRows(sheet, { limit: 2 });
    expect(rows).toHaveLength(2);
  });

  test('filters by statusColumn and statusValue', () => {
    const sheet = makeSheet(
      ['Name', 'Status'],
      [['Alice', 'active'], ['Bob', 'inactive'], ['Carol', 'active']],
    );
    const rows = getRows(sheet, { statusColumn: 'Status', statusValue: 'active' });
    expect(rows).toHaveLength(2);
    expect(rows.map(r => r.Name)).toEqual(['Alice', 'Carol']);
  });

  test('returns all rows when filter column is not present', () => {
    const sheet = makeSheet(
      ['Name'],
      [['Alice'], ['Bob']],
    );
    const rows = getRows(sheet, { statusColumn: 'Status', statusValue: 'active' });
    expect(rows).toHaveLength(0);
  });

  test('returns empty array for sheet with only headers', () => {
    const sheet = {
      getLastColumn: jest.fn(() => 2),
      getLastRow: jest.fn(() => 1),
      getRange: jest.fn(),
    };
    expect(getRows(sheet)).toEqual([]);
  });
});
