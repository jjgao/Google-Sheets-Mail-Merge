/**
 * Template Parser
 * Handles {{placeholder}} substitution in text strings.
 */

/**
 * Replace all {{Column Header}} placeholders in a string with values from a record.
 * Placeholders that have no matching key in the record are replaced with empty string.
 * @param {string} text - Text containing {{placeholder}} tokens
 * @param {Object} record - Map of column header → value
 * @returns {string} Text with all placeholders substituted
 */
function substitutePlaceholders(text, record) {
  if (!text) return text;

  return text.replace(/\{\{([^}]+)\}\}/g, function(match, key) {
    const value = record[key.trim()];
    if (value === undefined || value === null) return '';
    return value.toString();
  });
}
