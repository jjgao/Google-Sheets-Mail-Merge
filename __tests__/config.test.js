describe('CONFIG_KEYS', () => {
  test('defines all expected keys', () => {
    const expected = [
      'TEMPLATE_DOC_ID',
      'TEMPLATE_SLIDES_ID',
      'OUTPUT_FOLDER_ID',
    ];
    expected.forEach(key => {
      expect(CONFIG_KEYS[key]).toBe(key);
    });
  });
});

describe('validateConfig', () => {
  let originalGetConfig;

  beforeEach(() => {
    originalGetConfig = global.getConfig;
  });

  afterEach(() => {
    global.getConfig = originalGetConfig;
  });

  test('invalid when no config is set', () => {
    global.getConfig = jest.fn(() => null);
    const result = validateConfig();
    expect(result.isValid).toBe(false);
    expect(result.missing.length).toBeGreaterThan(0);
  });

  test('valid when doc template and output folder are set', () => {
    global.getConfig = jest.fn((key) => {
      if (key === CONFIG_KEYS.TEMPLATE_DOC_ID) return 'doc-id-123';
      if (key === CONFIG_KEYS.OUTPUT_FOLDER_ID) return 'folder-id-456';
      return null;
    });
    const result = validateConfig();
    expect(result.isValid).toBe(true);
    expect(result.missing).toHaveLength(0);
  });

  test('valid when slides template and output folder are set', () => {
    global.getConfig = jest.fn((key) => {
      if (key === CONFIG_KEYS.TEMPLATE_SLIDES_ID) return 'slides-id-123';
      if (key === CONFIG_KEYS.OUTPUT_FOLDER_ID) return 'folder-id-456';
      return null;
    });
    const result = validateConfig();
    expect(result.isValid).toBe(true);
    expect(result.missing).toHaveLength(0);
  });

  test('invalid when output folder is missing', () => {
    global.getConfig = jest.fn((key) => {
      if (key === CONFIG_KEYS.TEMPLATE_DOC_ID) return 'doc-id-123';
      return null;
    });
    const result = validateConfig();
    expect(result.isValid).toBe(false);
    expect(result.missing).toContain(CONFIG_KEYS.OUTPUT_FOLDER_ID);
  });

  test('invalid when template is missing but folder is set', () => {
    global.getConfig = jest.fn((key) => {
      if (key === CONFIG_KEYS.OUTPUT_FOLDER_ID) return 'folder-id-456';
      return null;
    });
    const result = validateConfig();
    expect(result.isValid).toBe(false);
    expect(result.missing).toContain(CONFIG_KEYS.TEMPLATE_DOC_ID);
    expect(result.missing).toContain(CONFIG_KEYS.TEMPLATE_SLIDES_ID);
  });

  test('returns missing as an array', () => {
    global.getConfig = jest.fn(() => null);
    const result = validateConfig();
    expect(Array.isArray(result.missing)).toBe(true);
  });
});
