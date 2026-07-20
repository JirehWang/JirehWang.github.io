(function(root, factory) {
  const api = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.PrayerPptExport = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(root) {

  const SLIDE_WIDTH = 13.333;
  const SLIDE_HEIGHT = 7.5;
  const slideX = percent => (Number(percent) / 100) * SLIDE_WIDTH;
  const slideY = percent => (Number(percent) / 100) * SLIDE_HEIGHT;

  function exportPrayerPPTX(options = {}) {
    const PptxGenJSClass = options.PptxGenJS || root.PptxGenJS;
    const getDeckEntriesFn = options.getDeckEntries || root.getDeckEntries;
    const layoutState = options.layoutState || root.worshipLayoutState;
    const production = options.production || root.PrayerSlideProduction;
    const model = options.model || root.model;
    const backgroundColor = options.backgroundColor || root.backgroundColor;
    const backgroundImage = options.backgroundImage || root.backgroundImage;
    const serviceDate = options.serviceDate || (document.getElementById('service-date') && document.getElementById('service-date').value) || '';

    const outputScale = layoutState && layoutState.outputScale || {};
    const textScale = (Number(outputScale.text) || 100) / 100;
    const scaledFont = value => Number(value) * textScale;

    if (!PptxGenJSClass) throw new Error('找不到 PptxGenJS 簡報庫');
    if (!getDeckEntriesFn) throw new Error('找不到 getDeckEntries 函式');
    if (!layoutState) throw new Error('找不到 layoutState');

    const pptx = new PptxGenJSClass();
    pptx.layout = 'LAYOUT_WIDE';

    const deck = getDeckEntriesFn();
    if (!deck || !deck.length) {
      throw new Error('沒有可匯出的投影片');
    }

    const bgFill = (backgroundColor || '#111111').replace('#', '');

    deck.forEach(entry => {
      const slide = pptx.addSlide();

      // Set background
      if (backgroundImage) {
        slide.background = { data: backgroundImage };
      } else {
        slide.background = { fill: bgFill };
      }

      // Get resolved layout
      const layout = production.layoutForPage(layoutState, entry);

      // 1. Add Title
      if (entry.title) {
        slide.addText(entry.title, {
          x: slideX(layout.titleX),
          y: slideY(layout.titleY),
          w: slideX(layout.titleW),
          h: slideY(layout.titleH),
          fontSize: scaledFont(layout.titleSize),
          fontFace: 'Microsoft JhengHei',
          color: (layout.titleColor || '#FFFFFF').replace('#', ''),
          align: layout.titleAlign || 'center',
          valign: 'middle',
          bold: true,
          margin: 0
        });
      }

      // 2. Add Content
      if (entry.body) {
        const wrappedBody = production.wrapTextForBox(entry.body, {
          fontSize: scaledFont(layout.contentSize),
          boxWidth: layout.contentW
        });

        slide.addText(wrappedBody, {
          x: slideX(layout.contentX),
          y: slideY(layout.contentY),
          w: slideX(layout.contentW),
          h: slideY(layout.contentH),
          fontSize: scaledFont(layout.contentSize),
          fontFace: 'Microsoft JhengHei',
          color: (layout.contentColor || '#E0E0E0').replace('#', ''),
          align: layout.contentAlign || 'left',
          valign: 'top',
          lineSpacing: scaledFont(layout.contentSize) * (layout.lineSpacing || 1.5),
          margin: 0
        });
      }
    });

    const dateStr = serviceDate ? `_${serviceDate}` : '';
    pptx.writeFile({ fileName: `林口教會_禱告會${dateStr}.pptx` });
  }

  return {
    exportPrayerPPTX
  };
});
