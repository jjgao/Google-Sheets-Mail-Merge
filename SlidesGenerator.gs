/**
 * Slides Generator
 * Clones a Google Slides template, tiles records into a grid layout,
 * substitutes {{placeholder}} fields, and saves the output to Drive.
 *
 * Template convention:
 *   The template Slides file must have exactly one slide (the label design).
 *   All shapes/text boxes on that slide define one label cell. The generator
 *   computes the bounding box of all elements to determine label dimensions,
 *   then calculates how many labels fit per slide (cols × rows). It appends
 *   one output slide per batch of records, leaving the original template
 *   slide untouched until the end (when it is removed).
 */

/**
 * Compute the bounding box that encloses all page elements on a slide.
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @returns {{ left: number, top: number, width: number, height: number }}
 */
function getLabelBoundingBox(slide) {
  const elements = slide.getPageElements();
  if (elements.length === 0) {
    return { left: 0, top: 0, width: 72, height: 72 };
  }

  let minLeft = Infinity;
  let minTop = Infinity;
  let maxRight = -Infinity;
  let maxBottom = -Infinity;

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    const left = el.getLeft();
    const top = el.getTop();
    const right = left + el.getWidth();
    const bottom = top + el.getHeight();

    if (left < minLeft) minLeft = left;
    if (top < minTop) minTop = top;
    if (right > maxRight) maxRight = right;
    if (bottom > maxBottom) maxBottom = bottom;
  }

  return {
    left: minLeft,
    top: minTop,
    width: maxRight - minLeft,
    height: maxBottom - minTop,
  };
}

/**
 * Detect the grid dimensions for a Slides template.
 * @param {GoogleAppsScript.Slides.Presentation} presentation
 * @returns {{ cols: number, rows: number, cellWidth: number, cellHeight: number, originLeft: number, originTop: number }}
 */
function detectSlidesGrid(presentation) {
  const slide = presentation.getSlides()[0];
  const slideWidth = presentation.getPageWidth();
  const slideHeight = presentation.getPageHeight();
  const bbox = getLabelBoundingBox(slide);

  const cols = Math.max(1, Math.floor(slideWidth / bbox.width));
  const rows = Math.max(1, Math.floor(slideHeight / bbox.height));

  return {
    cols: cols,
    rows: rows,
    cellWidth: bbox.width,
    cellHeight: bbox.height,
    originLeft: bbox.left,
    originTop: bbox.top,
  };
}

/**
 * Substitute placeholders in all text content of a page element.
 * @param {GoogleAppsScript.Slides.PageElement} element
 * @param {Object} record
 */
function substituteInSlideElement(element, record) {
  try {
    const shape = element.asShape ? element.asShape() : null;
    if (shape && shape.getText) {
      const textRange = shape.getText();
      const current = textRange.asRenderedString ? textRange.asRenderedString() : '';
      const replaced = substitutePlaceholders(current, record);
      if (replaced !== current) {
        textRange.setText(replaced);
      }
    }
  } catch (e) {
    Logger.log('substituteInSlideElement: ' + e.message);
  }
}

/**
 * Clear text content of a page element (for blank remainder cells).
 * @param {GoogleAppsScript.Slides.PageElement} element
 */
function clearSlideElement(element) {
  try {
    const shape = element.asShape ? element.asShape() : null;
    if (shape && shape.getText) {
      shape.getText().setText('');
    }
  } catch (e) { /* ignore */ }
}

/**
 * Populate one output slide with up to (cols × rows) records.
 * The slide is a fresh copy of the template (elements are in cell 0,0 position).
 * @param {GoogleAppsScript.Slides.Slide} slide - Fresh copy of template slide
 * @param {Object[]} records - All records
 * @param {number} startIndex - Index of first record for this slide
 * @param {{ cols, rows, cellWidth, cellHeight, originLeft, originTop }} grid
 */
function populateOutputSlide(slide, records, startIndex, grid) {
  // The slide has template elements positioned at (originLeft, originTop).
  // We'll use them for cell (0,0) and duplicate them for other cells.
  const originalElements = slide.getPageElements();

  // Pre-capture original element positions and text (before any modification)
  const origData = [];
  for (let i = 0; i < originalElements.length; i++) {
    const el = originalElements[i];
    let text = '';
    try {
      const shape = el.asShape ? el.asShape() : null;
      if (shape && shape.getText) {
        text = shape.getText().asRenderedString ? shape.getText().asRenderedString() : '';
      }
    } catch (e) { /* ignore */ }
    origData.push({
      left: el.getLeft(),
      top: el.getTop(),
      text: text,
    });
  }

  let recordIndex = startIndex;

  for (let row = 0; row < grid.rows; row++) {
    for (let col = 0; col < grid.cols; col++) {
      const offsetLeft = col * grid.cellWidth;
      const offsetTop = row * grid.cellHeight;
      const record = recordIndex < records.length ? records[recordIndex] : null;
      recordIndex++;

      if (row === 0 && col === 0) {
        // Use original elements in-place; they are already at (originLeft, originTop)
        for (let i = 0; i < originalElements.length; i++) {
          if (record) {
            substituteInSlideElement(originalElements[i], record);
          } else {
            clearSlideElement(originalElements[i]);
          }
        }
      } else {
        // Duplicate original elements and reposition for this cell
        for (let i = 0; i < originalElements.length; i++) {
          const duplicate = originalElements[i].duplicate();
          duplicate.setLeft(origData[i].left + offsetLeft);
          duplicate.setTop(origData[i].top + offsetTop);

          // Reset text to template value before substituting
          try {
            const shape = duplicate.asShape ? duplicate.asShape() : null;
            if (shape && shape.getText) {
              shape.getText().setText(origData[i].text);
            }
          } catch (e) { /* ignore */ }

          if (record) {
            substituteInSlideElement(duplicate, record);
          } else {
            clearSlideElement(duplicate);
          }
        }
      }
    }
  }
}

/**
 * Generate a tiled Google Slides file from a template and a list of records.
 * @param {Object[]} records - Array of record objects from SheetReader
 * @param {string} templateSlidesId - Google Slides file ID of the template
 * @param {string} outputFolderId - Drive folder ID for output
 * @returns {string} URL of the generated output file
 */
function generateSlides(records, templateSlidesId, outputFolderId) {
  const outputName = 'Labels - ' + Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const templateFile = DriveApp.getFileById(templateSlidesId);
  const outputFile = templateFile.makeCopy(outputName, outputFolder);

  const presentation = SlidesApp.openById(outputFile.getId());
  const grid = detectSlidesGrid(presentation);
  const cellsPerSlide = grid.cols * grid.rows;

  // Keep the template slide untouched; append output slides from it
  const templateSlide = presentation.getSlides()[0];
  const numSlides = Math.ceil(records.length / cellsPerSlide) || 1;

  for (let s = 0; s < numSlides; s++) {
    // appendSlide(templateSlide) creates a fresh copy of the template each time
    const outputSlide = presentation.appendSlide(templateSlide);
    populateOutputSlide(outputSlide, records, s * cellsPerSlide, grid);
  }

  // Remove the original template slide (now at index 0)
  templateSlide.remove();

  return outputFile.getUrl();
}
