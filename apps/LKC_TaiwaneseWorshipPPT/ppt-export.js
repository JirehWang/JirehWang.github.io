(function(root, factory) {
  const api = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipPptExport = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(root) {

  function cleanParagraphProperties(xmlString) {
    return xmlString.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (pMatch, pContent) => {
      let firstPPr = null;
      const cleanedContent = pContent.replace(/<a:pPr([\s\S]*?)<\/a:pPr>/g, (pprMatch) => {
        if (!firstPPr) {
          firstPPr = pprMatch;
          return pprMatch;
        } else {
          return '';
        }
      });
      return `<a:p>${cleanedContent}</a:p>`;
    });
  }

  function exportWorshipPPTX(options = {}) {
    const PptxGenJSClass = options.PptxGenJS || root.PptxGenJS;
    const getDeckEntriesFn = options.getDeckEntries || root.getDeckEntries;
    const layoutState = options.layoutState || root.worshipLayoutState;
    const production = options.production || root.TaiwaneseWorshipSlideProduction;
    const model = options.model || root.model;
    const backgroundColor = options.backgroundColor || root.backgroundColor;
    const backgroundImage = options.backgroundImage || root.backgroundImage;
    const serviceDate = options.serviceDate || (document.getElementById('service-date') && document.getElementById('service-date').value) || '';

    if (!PptxGenJSClass) throw new Error('找不到 PptxGenJS 簡報庫');
    if (!getDeckEntriesFn) throw new Error('找不到 getDeckEntries 函式');
    if (!layoutState) throw new Error('找不到 layoutState');

    const pptx = new PptxGenJSClass();
    pptx.layout = 'LAYOUT_16x9';

    const deck = getDeckEntriesFn();
    if (!deck || !deck.length) {
      throw new Error('沒有可匯出的投影片');
    }

    const bgFill = (backgroundColor || '#ffffff').replace('#', '');
    
    // Resolve rect shape type
    let rectShapeType = null;
    if (pptx.ShapeType && typeof pptx.ShapeType.rect !== 'undefined') {
      rectShapeType = pptx.ShapeType.rect;
    } else if (pptx.shapes && typeof pptx.shapes.RECTANGLE !== 'undefined') {
      rectShapeType = pptx.shapes.RECTANGLE;
    } else {
      rectShapeType = 'rect';
    }

    deck.forEach(entry => {
      const slide = pptx.addSlide();

      // 1. Background
      if (backgroundImage) {
        slide.background = { data: backgroundImage };
      } else {
        slide.background = { fill: bgFill };
      }

      // 2. White Overlay for hymns
      const hasHymnWhiteOverlay = (root.hymnOpacitySectionIds || []).includes(entry.sectionId);
      if (hasHymnWhiteOverlay && model && model[entry.sectionId]) {
        const opacityVal = model[entry.sectionId].opacity || 60;
        const transparency = 100 - opacityVal; // 60% opacity -> 40% transparency
        slide.addShape(rectShapeType, {
          x: 0,
          y: 0,
          w: 10,
          h: 5.625,
          fill: { color: 'FFFFFF', transparency: transparency }
        });
      }

      // 3. Layout Parameters
      const params = production.layoutForPage(layoutState, entry);

      const titleColor = (params.titleColor || '#111111').replace('#', '');
      const contentColor = (params.contentColor || '#111111').replace('#', '');

      // 4. Render Slide Content by kind
      if (entry.kind === 'ppt-import') {
        const finalObjects = getImportedSlideObjects(entry, params, production);
        finalObjects.forEach(obj => {
          if (obj.type === 'image' && obj.src) {
            slide.addImage({
              data: obj.src,
              x: (obj.x / 100) * 10,
              y: (obj.y / 100) * 5.625,
              w: (obj.w / 100) * 10,
              h: (obj.h / 100) * 5.625
            });
          } else if (obj.type === 'text' && Array.isArray(obj.runs)) {
            const runsArray = obj.runs.map(run => ({
              text: run.text,
              options: {
                fontSize: run.fontSize,
                color: (run.color || obj.color || '#000000').replace('#', ''),
                bold: run.bold,
                italic: run.italic,
                underline: run.underline,
                fontFace: run.fontFamily || obj.fontFamily
              }
            }));
            const verticalAlignMap = { start: 'top', center: 'middle', end: 'bottom' };
            slide.addText(runsArray, {
              x: (obj.x / 100) * 10,
              y: (obj.y / 100) * 5.625,
              w: (obj.w / 100) * 10,
              h: (obj.h / 100) * 5.625,
              align: obj.align || 'left',
              valign: verticalAlignMap[obj.verticalAlign] || 'top',
              lineSpacing: obj.lineSpacing,
              margin: 0
            });
          }
        });
      } else if (entry.kind === 'cover') {
        const [year, month, day] = serviceDate ? serviceDate.split('-') : [];
        const formattedDate = serviceDate ? `主後${year}年${month}月${day}日` : '';
        // Title
        slide.addText('台語主日禮拜', {
          x: (params.titleX / 100) * 10,
          y: (params.titleY / 100) * 5.625,
          w: (params.titleW / 100) * 10,
          h: (params.titleH / 100) * 5.625,
          fontSize: params.titleSize || 60,
          color: titleColor,
          fontFace: 'Microsoft JhengHei',
          align: params.titleAlign || 'center',
          valign: 'middle',
          bold: true,
          margin: 0
        });
        // Date
        slide.addText(formattedDate, {
          x: (params.contentX / 100) * 10,
          y: (params.contentY / 100) * 5.625,
          w: (params.contentW / 100) * 10,
          h: (params.contentH / 100) * 5.625,
          fontSize: params.contentSize || 48,
          color: contentColor,
          fontFace: 'Microsoft JhengHei',
          align: params.contentAlign || 'center',
          valign: 'top',
          margin: 0
        });
      } else if (entry.kind === 'praise-title') {
        slide.addText('讚美', {
          x: (params.titleX / 100) * 10,
          y: (params.titleY / 100) * 5.625,
          w: (params.titleW / 100) * 10,
          h: (params.titleH / 100) * 5.625,
          fontSize: params.titleSize || 60,
          color: titleColor,
          fontFace: 'Microsoft JhengHei',
          align: params.titleAlign || 'center',
          valign: 'middle',
          bold: true,
          margin: 0
        });
        const kicker = entry.kicker || (model && model[entry.sectionId] && model[entry.sectionId].kicker) || '';
        const titleText = entry.title || (model && model[entry.sectionId] && model[entry.sectionId].title) || '';
        const praiseContent = [titleText, kicker].filter(Boolean).join('\n\n');
        slide.addText(praiseContent, {
          x: (params.contentX / 100) * 10,
          y: (params.contentY / 100) * 5.625,
          w: (params.contentW / 100) * 10,
          h: (params.contentH / 100) * 5.625,
          fontSize: params.contentSize || 48,
          color: contentColor,
          fontFace: 'Microsoft JhengHei',
          align: params.contentAlign || 'center',
          valign: 'top',
          margin: 0
        });
      } else if (entry.kind === 'praise-lyrics') {
        slide.addText(entry.body || '', {
          x: (params.contentX / 100) * 10,
          y: (params.contentY / 100) * 5.625,
          w: (params.contentW / 100) * 10,
          h: (params.contentH / 100) * 5.625,
          fontSize: params.contentSize || 48,
          color: contentColor,
          fontFace: 'Microsoft JhengHei',
          align: params.contentAlign || 'center',
          valign: 'middle',
          lineSpacing: params.lineSpacing ? Math.round(params.contentSize * params.lineSpacing) : undefined,
          margin: 0
        });
      } else if (entry.kind === 'score') {
        const titleText = entry.title || (model && model[entry.sectionId] && model[entry.sectionId].title) || entry.sectionLabel || '';
        const kicker = entry.kicker || (model && model[entry.sectionId] && model[entry.sectionId].kicker) || '';
        // Title
        slide.addText(titleText, {
          x: (params.titleX / 100) * 10,
          y: (params.titleY / 100) * 5.625,
          w: (params.titleW / 100) * 10,
          h: (params.titleH / 100) * 5.625,
          fontSize: params.titleSize || 60,
          color: titleColor,
          fontFace: 'Microsoft JhengHei',
          align: params.titleAlign || 'center',
          valign: 'middle',
          bold: true,
          margin: 0
        });
        // Kicker/Sub
        if (kicker) {
          slide.addText(kicker, {
            x: (params.contentX / 100) * 10,
            y: ((params.contentY || 24) / 100) * 5.625,
            w: (params.contentW / 100) * 10,
            h: 1.0,
            fontSize: params.contentSize || 48,
            color: contentColor,
            fontFace: 'Microsoft JhengHei',
            align: params.contentAlign || 'center',
            valign: 'top',
            margin: 0
          });
        }
        // Score slot dashed box
        slide.addShape(rectShapeType, {
          x: 1.0,
          y: 3.2,
          w: 8.0,
          h: 2.0,
          fill: { color: 'FFFFFF', transparency: 30 },
          line: { color: '999999', width: 1, dashType: 'dash' }
        });
      } else {
        // Standard page: title and content
        const showTitle = entry.showTitle !== false;
        const titleText = entry.title || (model && model[entry.sectionId] && model[entry.sectionId].title) || entry.sectionLabel || '';
        
        if (showTitle && titleText) {
          slide.addText(titleText, {
            x: (params.titleX / 100) * 10,
            y: (params.titleY / 100) * 5.625,
            w: (params.titleW / 100) * 10,
            h: (params.titleH / 100) * 5.625,
            fontSize: params.titleSize || 60,
            color: titleColor,
            fontFace: 'Microsoft JhengHei',
            align: params.titleAlign || 'center',
            valign: 'middle',
            bold: true,
            margin: 0
          });
        }

        // Subtitles mapping
        const subtitles = { '會前領唱':'請準備心今天的禮拜', '靜默一分鐘':'請將手機關機或靜音', '後奏':'請後奏結束後再起身或交談', '平安禮':'請兄弟姊妹互相行平安禮' };
        const defaultBody = entry.kicker || subtitles[entry.sectionLabel] || '';
        const bodyText = entry.body || defaultBody;

        if (bodyText) {
          slide.addText(bodyText, {
            x: (params.contentX / 100) * 10,
            y: (params.contentY / 100) * 5.625,
            w: (params.contentW / 100) * 10,
            h: (params.contentH / 100) * 5.625,
            fontSize: params.contentSize || 48,
            color: contentColor,
            fontFace: 'Microsoft JhengHei',
            align: params.contentAlign || (entry.kind === 'section' ? 'center' : 'left'),
            valign: (entry.kind === 'section' ? 'top' : 'top'),
            lineSpacing: params.lineSpacing ? Math.round(params.contentSize * params.lineSpacing) : undefined,
            margin: 0
          });
        }
      }
    });

    const fileDate = serviceDate || new Date().toISOString().split('T')[0];
    const fileName = `台語主日禮拜_${fileDate}.pptx`;

    if (typeof pptx.write === 'function' && typeof document !== 'undefined') {
      return pptx.write('blob').then(async (blob) => {
        const JSZipLib = options.JSZip || root.JSZip;
        if (JSZipLib) {
          const zip = await JSZipLib.loadAsync(blob);
          const files = Object.keys(zip.files);
          for (const name of files) {
            if (name.startsWith('ppt/slides/slide') && name.endsWith('.xml')) {
              const originalXml = await zip.file(name).async('text');
              const cleanedXml = cleanParagraphProperties(originalXml);
              zip.file(name, cleanedXml);
            }
          }
          const cleanedBlob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
          });
          const url = URL.createObjectURL(cleanedBlob);
          const a = document.createElement('a');
          a.href = url;
          a.download = fileName;
          document.body.appendChild(a);
          a.click();
          setTimeout(() => {
            URL.revokeObjectURL(url);
            document.body.removeChild(a);
          }, 100);
          return;
        }
        return pptx.writeFile({ fileName: fileName });
      });
    }

    return pptx.writeFile({ fileName: fileName });
  }

  function getImportedSlideObjects(page, params, production) {
    const objects = page.objects || [];
    const textObjects = objects.filter(obj => obj.type === 'text');
    
    const computeRoleBounds = (role) => {
      const roleObjects = textObjects.filter(obj => obj.role === role);
      if (!roleObjects.length) return null;
      const x = Math.min(...roleObjects.map(obj => obj.x));
      const y = Math.min(...roleObjects.map(obj => obj.y));
      const right = Math.max(...roleObjects.map(obj => obj.x + obj.w));
      const bottom = Math.max(...roleObjects.map(obj => obj.y + obj.h));
      return { x, y, w: right - x, h: bottom - y };
    };

    const titleBounds = computeRoleBounds('title');
    const contentBounds = computeRoleBounds('content');

    return objects.map(obj => {
      if (obj.type === 'image') {
        return { ...obj };
      }
      
      const prefix = obj.role === 'title' ? 'title' : 'content';
      const bounds = obj.role === 'title' ? titleBounds : contentBounds;
      if (!bounds || params[`${prefix}X`] == null) {
        return { ...obj };
      }
      
      const scaleX = Number(params[`${prefix}W`]) / Math.max(bounds.w, 0.01);
      const scaleY = Number(params[`${prefix}H`]) / Math.max(bounds.h, 0.01);

      const finalX = Number(params[`${prefix}X`]) + (obj.x - bounds.x) * scaleX;
      const finalY = Number(params[`${prefix}Y`]) + (obj.y - bounds.y) * scaleY;
      const finalW = obj.w * scaleX;
      const finalH = obj.h * scaleY;

      const mappedRuns = (obj.runs || []).map(run => {
        const sourceBaseSize = Number(obj.fontSize) || 18;
        const relativeSize = (Number(run.fontSize) || sourceBaseSize) / sourceBaseSize;
        const finalFontSize = (Number(params[`${prefix}Size`]) * relativeSize); 
        const finalColor = params[`${prefix}Color`] || run.color || obj.color || '#000000';

        return {
          ...run,
          fontSize: Number(finalFontSize.toFixed(1)),
          color: finalColor
        };
      });

      return {
        ...obj,
        x: finalX,
        y: finalY,
        w: finalW,
        h: finalH,
        align: params[`${prefix}Align`] || obj.align || 'left',
        runs: mappedRuns,
        lineSpacing: params.lineSpacing ? Math.round((Number(params[`${prefix}Size`]) || 18) * params.lineSpacing) : undefined
      };
    });
  }

  return {
    exportWorshipPPTX,
    getImportedSlideObjects,
    cleanParagraphProperties
  };
});
