const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.join(__dirname, '..');

function loadBulletinModel() {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayBulletin', 'js', 'bulletin.js'),
    'utf8'
  );
  const context = {
    CONFIG: { BANK_ACCOUNT: '', TW_GROUPS: [], SUNDAY_SCHOOL_CLASSES: [] },
    window: {},
    module: { exports: {} },
    exports: {}
  };
  vm.createContext(context);
  vm.runInContext(`${source}\nmodule.exports = BulletinModel;`, context);
  return context.module.exports;
}

function loadChurchApi() {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayBulletin', 'js', 'api.js'),
    'utf8'
  );
  const context = {
    CONFIG: { LKGROUP_GAS_URL: 'https://example.test/group', TW_GROUPS: ['恩友小組'] },
    window: {},
    formatYMD(date) {
      return date.toISOString().slice(0, 10);
    },
    debug() {},
    console,
    module: { exports: {} },
    exports: {}
  };
  vm.createContext(context);
  vm.runInContext(`${source}\nmodule.exports = ChurchAPI;`, context);
  return context.module.exports;
}

function loadBulletinExport() {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayBulletin', 'js', 'export.js'),
    'utf8'
  );
  const makeDocxType = name => class {
    constructor(props) {
      this.type = name;
      Object.assign(this, props);
    }
  };
  const context = {
    window: {
      docx: {
        TextRun: makeDocxType('TextRun'),
        Paragraph: makeDocxType('Paragraph'),
        TableCell: makeDocxType('TableCell'),
        TableRow: makeDocxType('TableRow'),
        Table: makeDocxType('Table'),
        AlignmentType: { LEFT: 'LEFT', CENTER: 'CENTER' },
        WidthType: { DXA: 'DXA' },
        BorderStyle: { NONE: 'NONE', SINGLE: 'SINGLE' }
      }
    },
    module: { exports: {} },
    exports: {}
  };
  vm.createContext(context);
  vm.runInContext(`${source}\nmodule.exports = BulletinExport;`, context);
  return context.module.exports;
}

test('Word page 1 export includes the uploaded choir lyrics', () => {
  const exporter = loadBulletinExport();
  const page = exporter._buildPage1({
    date: '2026-08-09',
    serviceType: '台華語',
    taiwanese: {
      presider: '', callToWorship: '', openingHymn: '', responsivePsalm: '',
      prayer1Note: '', scripture: '', choirSong: '讚美曲目', choirLyrics: '第一段\n第二段',
      sermonTitle: '', responseHymn: '', goldenVerse: '', goldenVerseText: '',
      offeringNote: '', doxologyHymn: '', bankAccount: ''
    },
    mandarin: {
      presider: '', goldenVerse: '', goldenVerseText: '', scripture: '', sermonTitle: '',
      worshipSongs: '', upcomingPreview: ''
    }
  });

  assert.match(JSON.stringify(page), /讚美曲目/);
  assert.match(JSON.stringify(page), /第一段/);
  assert.match(JSON.stringify(page), /第二段/);
});

test('small-group query errors are returned as failures instead of empty attendance', async () => {
  const api = loadChurchApi();
  api.callGAS = async (url, action) => {
    if (action === 'getGroups') return { success: true, groups: [{ name: '恩友小組' }] };
    throw new Error('小組週報服務暫時無法連線');
  };

  const result = await api.fetchSmallGroups('2026-08-09');

  assert.equal(result.success, false);
  assert.match(result.error, /小組週報服務暫時無法連線/);
  assert.equal(result.data, undefined);
});

test('a confirmed empty small-group report is marked as no meeting', async () => {
  const api = loadChurchApi();
  api.callGAS = async (url, action) => {
    if (action === 'getGroups') return { success: true, groups: [{ name: '恩友小組' }] };
    return { success: true, dateRange: '2026-08-02 ~ 2026-08-08', data: [] };
  };

  const result = await api.fetchSmallGroups('2026-08-09');

  assert.equal(result.success, true);
  assert.equal(result.data['恩友小組'].attendance, '本週無聚會');
});

test('missing current-date hymn and responsive-psalm values do not retain previous values', () => {
  const model = loadBulletinModel();
  model.init('2026-08-09');
  model.set('taiwanese.responsivePsalm', '12');
  model.set('taiwanese.openingHymn', '101');
  model.set('taiwanese.responseHymn', '202');
  model.set('taiwanese.doxologyHymn', '303');

  model.applyAPIData({
    calendar: {
      success: true,
      data: { taiwanese: { hymn: '', responsivePsalm: '' }, mandarin: null, upcoming: [] }
    }
  });

  const data = model.get();
  assert.equal(data.taiwanese.responsivePsalm, '無資料');
  assert.equal(data.taiwanese.openingHymn, '無資料');
  assert.equal(data.taiwanese.responseHymn, '無資料');
  assert.equal(data.taiwanese.doxologyHymn, '無資料');
});
