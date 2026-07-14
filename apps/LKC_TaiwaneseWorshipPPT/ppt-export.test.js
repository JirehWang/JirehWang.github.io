const test = require('node:test');
const assert = require('node:assert/strict');
const { exportWorshipPPTX, getImportedSlideObjects } = require('./ppt-export.js');

test('exports slides correctly with mock pptxgenjs', async () => {
  const slides = [];
  class MockPptx {
    constructor() {
      this.layout = '';
    }
    addSlide() {
      const slide = {
        background: null,
        shapes: [],
        images: [],
        texts: [],
        addShape(type, opts) {
          this.shapes.push({ type, opts });
        },
        addImage(opts) {
          this.images.push(opts);
        },
        addText(text, opts) {
          this.texts.push({ text, opts });
        }
      };
      slides.push(slide);
      return slide;
    }
    writeFile(opts) {
      this.writtenFileOpts = opts;
      return Promise.resolve();
    }
  }

  MockPptx.ShapeType = { rect: 'rect' };

  const mockDeck = [
    { kind: 'cover', sectionId: 'cover', sectionLabel: '台語主日禮拜' },
    { kind: 'score', sectionId: 'hymn-1', sectionLabel: '聖詩一', title: '聖詩第1首', kicker: '第一首' },
    { kind: 'ppt-import', sectionId: 'hymn-1', sectionLabel: '聖詩一', objects: [
      { type: 'image', src: 'data:image/png;base64,123', x: 10, y: 10, w: 20, h: 20 },
      { type: 'text', role: 'content', text: '歌詞內容', x: 15, y: 25, w: 70, h: 50, runs: [{ text: '歌詞內容', fontSize: 18 }] }
    ]}
  ];

  const mockLayoutState = { groups: {}, pageAssignments: {} };
  const mockProduction = {
    layoutForPage: () => ({
      titleSize: 60, titleX: 10, titleY: 6, titleW: 80, titleH: 16, titleAlign: 'center', titleColor: '#111111',
      contentSize: 48, contentX: 8, contentY: 24, contentW: 84, contentH: 68, contentAlign: 'left', contentColor: '#111111', lineSpacing: 1.5
    })
  };

  const mockModel = {
    'hymn-1': { opacity: 50 }
  };

  // Run export
  const exportOptions = {
    PptxGenJS: MockPptx,
    getDeckEntries: () => mockDeck,
    layoutState: mockLayoutState,
    production: mockProduction,
    model: mockModel,
    backgroundColor: '#ffffff',
    backgroundImage: '',
    serviceDate: '2026-07-14'
  };

  // Mock global/root variables
  globalThis.hymnOpacitySectionIds = ['hymn-1'];

  await exportWorshipPPTX(exportOptions);

  assert.equal(slides.length, 3);
  
  // Verify slide 1: cover
  assert.equal(slides[0].texts.length, 2);
  assert.equal(slides[0].texts[0].text, '台語主日禮拜');
  assert.equal(slides[0].texts[1].text, '主後2026年07月14日');

  // Verify slide 2: score (hymn-1) has white overlay shape
  assert.equal(slides[1].shapes.length, 2);
  assert.equal(slides[1].shapes[0].type, 'rect');
  assert.equal(slides[1].shapes[0].opts.fill.transparency, 50); // 100 - opacity(50) = 50
  assert.equal(slides[1].texts[0].text, '聖詩第1首');

  // Verify slide 3: ppt-import
  assert.equal(slides[2].images.length, 1);
  assert.equal(slides[2].images[0].data, 'data:image/png;base64,123');
  assert.equal(slides[2].texts.length, 1);
  assert.equal(slides[2].texts[0].text[0].text, '歌詞內容');
});
