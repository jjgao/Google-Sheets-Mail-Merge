describe('substitutePlaceholders', () => {
  test('replaces a single placeholder', () => {
    expect(substitutePlaceholders('Hello {{First Name}}', { 'First Name': 'Alice' }))
      .toBe('Hello Alice');
  });

  test('replaces multiple placeholders', () => {
    expect(substitutePlaceholders('{{First Name}} {{Last Name}}', { 'First Name': 'Alice', 'Last Name': 'Smith' }))
      .toBe('Alice Smith');
  });

  test('replaces the same placeholder multiple times', () => {
    expect(substitutePlaceholders('{{Name}} and {{Name}}', { Name: 'Bob' }))
      .toBe('Bob and Bob');
  });

  test('replaces missing key with empty string', () => {
    expect(substitutePlaceholders('Hello {{Missing}}', {}))
      .toBe('Hello ');
  });

  test('handles null record value as empty string', () => {
    expect(substitutePlaceholders('Hello {{Name}}', { Name: null }))
      .toBe('Hello ');
  });

  test('coerces numeric values to string', () => {
    expect(substitutePlaceholders('Count: {{N}}', { N: 42 }))
      .toBe('Count: 42');
  });

  test('returns text unchanged when no placeholders present', () => {
    expect(substitutePlaceholders('No placeholders here', { Name: 'Alice' }))
      .toBe('No placeholders here');
  });

  test('returns empty string for empty input', () => {
    expect(substitutePlaceholders('', { Name: 'Alice' })).toBe('');
  });

  test('returns null/undefined as-is', () => {
    expect(substitutePlaceholders(null, {})).toBeNull();
    expect(substitutePlaceholders(undefined, {})).toBeUndefined();
  });

  test('trims whitespace inside placeholder braces', () => {
    expect(substitutePlaceholders('{{ First Name }}', { 'First Name': 'Alice' }))
      .toBe('Alice');
  });
});
