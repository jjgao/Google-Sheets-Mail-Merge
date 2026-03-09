/**
 * Jest setup for Google Apps Script unit tests.
 *
 * GAS source files are loaded via require() using jest.gastransform.js,
 * which converts top-level function declarations and const/let/var declarations
 * to global.* assignments — mirroring the shared global scope GAS provides.
 *
 * Load order matters: each file must be loaded after its dependencies.
 */

// ---------------------------------------------------------------------------
// Mock GAS globals (must be defined before requiring source files)
// ---------------------------------------------------------------------------

function makeMockSheet() {
  return {
    getName: jest.fn(() => 'TestSheet'),
    getLastColumn: jest.fn(() => 0),
    getLastRow: jest.fn(() => 0),
    getDataRange: jest.fn(() => ({
      getValues: jest.fn(() => []),
      getDisplayValues: jest.fn(() => []),
    })),
    getRange: jest.fn(() => ({
      getValues: jest.fn(() => [[]]),
      getDisplayValues: jest.fn(() => [[]]),
      setValues: jest.fn(),
      setValue: jest.fn(),
      setFontWeight: jest.fn(),
      setBackground: jest.fn(),
      setFontColor: jest.fn(),
      setFontSize: jest.fn(),
      setWrap: jest.fn(),
      merge: jest.fn(),
      protect: jest.fn(() => ({ setDescription: jest.fn(), setWarningOnly: jest.fn() })),
    })),
    setFrozenRows: jest.fn(),
    setColumnWidth: jest.fn(),
    setRowHeight: jest.fn(),
    insertRowBefore: jest.fn(),
  };
}

global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(() => ({
    getSheetByName: jest.fn(() => null),
    getActiveSheet: jest.fn(() => makeMockSheet()),
    getSheets: jest.fn(() => []),
    insertSheet: jest.fn(() => makeMockSheet()),
    deleteSheet: jest.fn(),
    moveActiveSheet: jest.fn(),
    getUi: jest.fn(() => ({
      alert: jest.fn(),
      ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' },
    })),
    toast: jest.fn(),
  })),
  getUi: jest.fn(() => ({
    alert: jest.fn(),
    ButtonSet: { OK: 'OK' },
    createMenu: jest.fn(() => ({
      addItem: jest.fn().mockReturnThis(),
      addSeparator: jest.fn().mockReturnThis(),
      addToUi: jest.fn(),
    })),
  })),
  flush: jest.fn(),
};

global.DriveApp = {
  getFileById: jest.fn(() => ({
    getBlob: jest.fn(() => 'blob'),
    makeCopy: jest.fn(() => ({
      getId: jest.fn(() => 'copy-id'),
      getUrl: jest.fn(() => 'https://drive.google.com/file/copy-id'),
      setName: jest.fn(),
    })),
  })),
  getFolderById: jest.fn(() => ({
    getId: jest.fn(() => 'folder-id'),
  })),
  getFoldersByName: jest.fn(() => ({ hasNext: jest.fn(() => false), next: jest.fn() })),
  createFolder: jest.fn(() => ({ getId: jest.fn(() => 'test-folder-id') })),
};

global.DocumentApp = {
  openById: jest.fn(),
  create: jest.fn(),
  ElementType: {
    PARAGRAPH: 'PARAGRAPH',
    LIST_ITEM: 'LIST_ITEM',
    TABLE: 'TABLE',
    TABLE_ROW: 'TABLE_ROW',
    TABLE_CELL: 'TABLE_CELL',
  },
  ParagraphHeading: { HEADING1: 'H1', HEADING2: 'H2', HEADING3: 'H3', NORMAL: 'NORMAL' },
  HorizontalAlignment: { LEFT: 'LEFT', CENTER: 'CENTER', RIGHT: 'RIGHT', JUSTIFY: 'JUSTIFY' },
};

global.SlidesApp = {
  openById: jest.fn(),
  create: jest.fn(),
};

global.PropertiesService = {
  getDocumentProperties: jest.fn(() => ({
    getProperty: jest.fn(() => null),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
  })),
};

global.Utilities = {
  sleep: jest.fn(),
  formatDate: jest.fn((date, tz, fmt) => '2026-03-08'),
};

global.Logger = {
  log: jest.fn(),
};

global.Session = {
  getActiveUser: jest.fn(() => ({ getEmail: jest.fn(() => 'user@example.com') })),
  getEffectiveUser: jest.fn(() => ({ getEmail: jest.fn(() => 'user@example.com') })),
  getScriptTimeZone: jest.fn(() => 'America/New_York'),
};

// ---------------------------------------------------------------------------
// Load GAS source files (transform assigns top-level declarations to global.*)
// ---------------------------------------------------------------------------

require('./Config.gs');
require('./SheetReader.gs');
require('./TemplateParser.gs');
require('./DocumentGenerator.gs');
require('./SlidesGenerator.gs');
