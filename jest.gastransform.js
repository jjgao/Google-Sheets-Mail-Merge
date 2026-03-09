/**
 * Jest transform for Google Apps Script (.gs) files.
 *
 * GAS files define functions and variables in a shared global scope across
 * all files. This transform makes that behaviour work in Jest by converting:
 *   - Top-level function declarations  → global.name = function(...)
 *   - Top-level const/let/var decls    → global.name = ...
 *
 * "Top-level" is identified by a match at the start of a line (no indentation),
 * which reliably separates module-level declarations from those inside functions.
 */
module.exports = {
  process(sourceText) {
    let code = sourceText;

    // Top-level function declarations → global assignments
    code = code.replace(/^function\s+(\w+)\s*\(/gm, 'global.$1 = function(');

    // Top-level const / let / var declarations → global assignments
    code = code.replace(/^(?:const|let|var)\s+(\w+)(\s*=)/gm, 'global.$1$2');

    return { code: `${code}\nmodule.exports = {};` };
  },
};
