describe('getLabelBoundingBox', () => {
  function makeSlide(elements) {
    return { getPageElements: jest.fn(() => elements) };
  }

  function makeElement(left, top, width, height) {
    return {
      getLeft: jest.fn(() => left),
      getTop: jest.fn(() => top),
      getWidth: jest.fn(() => width),
      getHeight: jest.fn(() => height),
    };
  }

  test('returns default for empty slide', () => {
    const slide = makeSlide([]);
    expect(getLabelBoundingBox(slide)).toEqual({ left: 0, top: 0, width: 72, height: 72 });
  });

  test('returns element dimensions for single element', () => {
    const slide = makeSlide([makeElement(10, 20, 100, 50)]);
    expect(getLabelBoundingBox(slide)).toEqual({ left: 10, top: 20, width: 100, height: 50 });
  });

  test('returns bounding box of multiple elements', () => {
    const slide = makeSlide([
      makeElement(10, 20, 100, 50),  // right=110, bottom=70
      makeElement(50, 10, 80, 90),   // right=130, bottom=100
    ]);
    expect(getLabelBoundingBox(slide)).toEqual({ left: 10, top: 10, width: 120, height: 90 });
  });
});

describe('detectSlidesGrid', () => {
  function makePresentation({ slideWidth = 720, slideHeight = 540, bboxWidth = 180, bboxHeight = 135 } = {}) {
    const mockSlide = {
      getPageElements: jest.fn(() => [{
        getLeft: jest.fn(() => 0),
        getTop: jest.fn(() => 0),
        getWidth: jest.fn(() => bboxWidth),
        getHeight: jest.fn(() => bboxHeight),
      }]),
    };
    return {
      getSlides: jest.fn(() => [mockSlide]),
      getPageWidth: jest.fn(() => slideWidth),
      getPageHeight: jest.fn(() => slideHeight),
    };
  }

  test('computes cols and rows from slide and cell dimensions', () => {
    // slideWidth=720, bboxWidth=180 → cols=4; slideHeight=540, bboxHeight=135 → rows=4
    const pres = makePresentation({ slideWidth: 720, slideHeight: 540, bboxWidth: 180, bboxHeight: 135 });
    const grid = detectSlidesGrid(pres);
    expect(grid.cols).toBe(4);
    expect(grid.rows).toBe(4);
  });

  test('returns at least 1 col and 1 row', () => {
    const pres = makePresentation({ bboxWidth: 10000, bboxHeight: 10000 });
    const grid = detectSlidesGrid(pres);
    expect(grid.cols).toBeGreaterThanOrEqual(1);
    expect(grid.rows).toBeGreaterThanOrEqual(1);
  });

  test('includes cellWidth and cellHeight in result', () => {
    const pres = makePresentation({ bboxWidth: 180, bboxHeight: 135 });
    const grid = detectSlidesGrid(pres);
    expect(grid.cellWidth).toBe(180);
    expect(grid.cellHeight).toBe(135);
  });
});

describe('generateSlides', () => {
  function makeElement(text = '{{Name}}') {
    const textRange = {
      asRenderedString: jest.fn(() => text),
      setText: jest.fn(),
    };
    const shape = { getText: jest.fn(() => textRange) };
    const el = {
      getLeft: jest.fn(() => 0),
      getTop: jest.fn(() => 0),
      getWidth: jest.fn(() => 100),
      getHeight: jest.fn(() => 50),
      asShape: jest.fn(() => shape),
      duplicate: jest.fn(function() { return makeElement(text); }),
      setLeft: jest.fn(),
      setTop: jest.fn(),
    };
    return el;
  }

  function makeSlide(elements) {
    return {
      getPageElements: jest.fn(() => elements),
      remove: jest.fn(),
    };
  }

  function makePresentationMock(templateElements) {
    const templateSlide = makeSlide(templateElements);
    const outputSlides = [];

    return {
      getSlides: jest.fn(() => [templateSlide, ...outputSlides]),
      getPageWidth: jest.fn(() => 720),
      getPageHeight: jest.fn(() => 540),
      appendSlide: jest.fn(() => {
        const newSlide = makeSlide(templateElements.map(el => makeElement()));
        outputSlides.push(newSlide);
        return newSlide;
      }),
      saveAndClose: jest.fn(),
      _templateSlide: templateSlide,
    };
  }

  beforeEach(() => {
    const outputFileMock = {
      getId: jest.fn(() => 'output-slides-id'),
      getUrl: jest.fn(() => 'https://docs.google.com/presentation/output-slides-id'),
    };
    DriveApp.getFolderById.mockReturnValue({ getId: jest.fn(() => 'folder-id') });
    DriveApp.getFileById.mockReturnValue({
      makeCopy: jest.fn(() => outputFileMock),
    });

    const pres = makePresentationMock([makeElement('{{Name}}')]);
    SlidesApp.openById.mockReturnValue(pres);
  });

  test('returns a URL string', () => {
    const url = generateSlides([{ Name: 'Alice' }], 'template-slides-id', 'folder-id');
    expect(typeof url).toBe('string');
    expect(url).toContain('http');
  });

  test('calls makeCopy on the template file', () => {
    generateSlides([{ Name: 'Alice' }], 'template-slides-id', 'folder-id');
    expect(DriveApp.getFileById).toHaveBeenCalledWith('template-slides-id');
    const fileMock = DriveApp.getFileById.mock.results[0].value;
    expect(fileMock.makeCopy).toHaveBeenCalled();
  });

  test('output file name includes "Labels"', () => {
    generateSlides([{ Name: 'Alice' }], 'template-slides-id', 'folder-id');
    const fileMock = DriveApp.getFileById.mock.results[0].value;
    const copyName = fileMock.makeCopy.mock.calls[0][0];
    expect(copyName).toMatch(/Labels/);
  });

  test('opens the copied presentation', () => {
    generateSlides([{ Name: 'Alice' }], 'template-slides-id', 'folder-id');
    expect(SlidesApp.openById).toHaveBeenCalledWith('output-slides-id');
  });

  test('appends at least one output slide', () => {
    generateSlides([{ Name: 'Alice' }], 'template-slides-id', 'folder-id');
    const pres = SlidesApp.openById.mock.results[0].value;
    expect(pres.appendSlide).toHaveBeenCalled();
  });

  test('removes template slide', () => {
    generateSlides([{ Name: 'Alice' }], 'template-slides-id', 'folder-id');
    const pres = SlidesApp.openById.mock.results[0].value;
    const templateSlide = pres._templateSlide;
    expect(templateSlide.remove).toHaveBeenCalled();
  });
});
