# Google Sheets Mail Merge — Claude Code Guide

## Project Overview

A Google Apps Script add-on that merges data from a Google Sheet into a grid/tiled layout — generating label sheets, name badges, event passes, or any repeated per-record layout — and saves the output to Google Drive.

Each row in the sheet is one record. The user designs a single label/badge in a Google Doc or Google Slides template using `{{placeholder}}` syntax. The tool auto-detects the grid layout from template dimensions, tiles all records across pages/slides, and saves the output to a configured Drive folder.

**Current version:** 1.0.0

## Tech Stack

- **Platform:** Google Apps Script (V8 runtime, JavaScript ES6)
- **Deployment:** [Clasp CLI](https://github.com/google/clasp) — `.clasp.json` is gitignored; copy from `.clasp.json.example` for local use
- **Testing:** Jest with a custom `.gs` transform (`jest.gastransform.js`) that rewrites GAS files to use `global.*` assignments
- **CI/CD:** GitHub Actions — CI on every push/PR, CD (`clasp push`) on merge to `main` using `CLASP_TOKEN` and `SCRIPT_ID` repo secrets

## File Structure

| File | Responsibility |
|------|---------------|
| `Code.gs` | App entry point, version info, `initializeApp()` |
| `Config.gs` | Config key constants, read/write from Config sheet, validation |
| `SheetReader.gs` | Read headers and rows from active sheet, row filtering |
| `TemplateParser.gs` | `{{placeholder}}` substitution in text strings |
| `DocumentGenerator.gs` | Clone Doc template, detect grid, tile records, save to Drive |
| `SlidesGenerator.gs` | Clone Slides template, detect grid, tile records, save to Drive |
| `UI.gs` | Custom menu, config dialog, generate action, progress toasts |
| `appsscript.json` | OAuth scopes and runtime config |

## Development Workflow

```bash
# Push local changes to Apps Script
clasp push

# Pull from Apps Script (after editing in browser editor)
clasp pull

# Open the project in browser
clasp open
```

Changes to `.gs` files are pushed directly — no build step needed.

## Contribution Workflow (Features & Fixes)

For every code change — bug fix, feature, or refactor — follow this flow:

**1. Create a GitHub issue**
```bash
gh issue create --repo jjgao/Google-Sheets-Mail-Merge --title "..." --label "bug|enhancement" --body "..."
```

**2. Create a branch named after the issue**
```bash
git checkout -b issue-<number>-short-description
# e.g. git checkout -b issue-5-fix-grid-detection
```

**3. Make changes, then open a pull request**
```bash
git add <files>
git commit -m "Fix: short description (#<issue-number>)"
gh pr create --title "..." --body "..." --head issue-<number>-... --base main
```

The PR body should reference the issue (`Closes #<number>`) so it auto-closes on merge.

> **IMPORTANT:** Never merge a pull request unless the user explicitly asks to merge it. Always stop after opening the PR and wait for instruction.

## Code Conventions

- **Functions:** camelCase
- **Constants:** UPPER_SNAKE_CASE
- **Private properties:** underscore prefix
- **UI wrapper functions:** suffix with `UI` (e.g., `runGenerateUI`)
- All config keys go through `CONFIG_KEYS` object in `Config.gs`
- Validation functions return `{ isValid: boolean, error: string }` or `{ isValid: boolean, missing: [] }`

## Key Patterns

**Configuration:** All settings stored in a "Config" Google Sheet tab (not Script Properties). Access via `getConfig(CONFIG_KEYS.KEY_NAME)`.

**Template syntax:** `{{Column Header}}` — case-sensitive, matches sheet column header exactly. Missing placeholders are replaced with empty string.

**Grid detection:** The template defines one label cell. `detectGrid()` in each generator computes `cols = floor(pageWidth / cellWidth)`, `rows = floor(pageHeight / cellHeight)` from page/slide dimensions and the label bounding box.

**Output naming:** `Labels - YYYY-MM-DD` using `Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")`.

**Generate flow:**
1. `UI.runGenerate()` reads config + sheet data
2. Calls `DocumentGenerator.generateDocument()` or `SlidesGenerator.generateSlides()` depending on which template is configured
3. Shows toast with link to the output file

## System Sheets (Do Not Rename)

- `Config` — configuration key-value store

## Running Tests

**Unit tests (Jest):**
```bash
npm install
npm test
```

**System test (in Apps Script Editor):** run `initializeApp()` to set up the Config sheet.
