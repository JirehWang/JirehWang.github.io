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
    CONFIG: {
      BANK_ACCOUNT: '',
      TW_GROUPS: [],
      SUNDAY_SCHOOL_CLASSES: []
    },
    window: {},
    module: { exports: {} },
    exports: {}
  };
  vm.createContext(context);
  vm.runInContext(`${source}\nmodule.exports = BulletinModel;`, context);
  return context.module.exports;
}

function loadApp(model, responseData) {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayBulletin', 'js', 'app.js'),
    'utf8'
  );
  const context = {
    window: {},
    document: {
      addEventListener() {},
      querySelector() { return null; },
      querySelectorAll() { return []; }
    },
    BulletinModel: model,
    ChurchAPI: {},
    CONFIG: { GAS_SYNC_URL: 'https://example.test/sync' },
    fetch: async () => ({ ok: true, async json() { return responseData; } }),
    console,
    module: { exports: {} },
    exports: {}
  };
  vm.createContext(context);
  vm.runInContext(`${source}\nmodule.exports = { App, normalizeUploadedPraise, normalizeUploadedReports };`, context);
  return context.module.exports;
}

test('the weekly bulletin model exposes a choir-lyrics field for the uploaded praise record', () => {
  const model = loadBulletinModel();
  model.init('2026-08-09');

  assert.equal(model.get().taiwanese.choirLyrics, '');

  const html = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayBulletin', 'index.html'),
    'utf8'
  );
  assert.match(html, /data-field="taiwanese\.choirLyrics"/);
});

test('loading uploaded praise fills both the choir song and lyrics fields', async () => {
  const model = loadBulletinModel();
  model.init('2026-08-09');
  const { App, normalizeUploadedPraise } = loadApp(model, {
    success: true,
    data: {
      title: '新的事將要成就',
      kicker: '聖歌隊',
      lyrics: '第一段\n\n第二段'
    }
  });

  assert.deepEqual(JSON.parse(JSON.stringify(normalizeUploadedPraise({ title: ' 歌名 ', lyrics: ' 歌詞 ' }))), {
    title: '歌名',
    lyrics: '歌詞'
  });

  App._els = { bulletinDate: { value: '2026-08-09' } };
  App.syncFormFromModel = () => {};
  const result = await App.loadUploadedChoirSong({ silent: true });

  assert.deepEqual(JSON.parse(JSON.stringify(result)), { failed: [] });
  assert.equal(model.get().taiwanese.choirSong, '新的事將要成就');
  assert.equal(model.get().taiwanese.choirLyrics, '第一段\n\n第二段');
});

test('loading uploaded reports replaces all weekly bulletin report fields instead of leaving stale entries', async () => {
  const model = loadBulletinModel();
  model.init('2026-08-09');
  model.set('announcements.4', '舊的本會消息');
  model.set('churchNews.4', '舊的教界消息');
  model.set('prayer.homeRest', '舊的代禱');

  const { App, normalizeUploadedReports } = loadApp(model, {
    success: true,
    data: {
      announcements: ['本會消息一'],
      churchNews: ['教界消息一', '教界消息二'],
      prayer: { hospital: '住院代禱' }
    }
  });

  assert.deepEqual(JSON.parse(JSON.stringify(normalizeUploadedReports({ announcements: ['消息'] }))), {
    announcements: ['消息', '', '', '', '', '', '', '', '', ''],
    churchNews: ['', '', '', '', '', '', '', '', '', ''],
    prayer: { homeRest: '', hospital: '', other: '' }
  });

  App._els = { bulletinDate: { value: '2026-08-09' } };
  App.syncFormFromModel = () => {};
  const result = await App.loadUploadedReports({ silent: true });

  assert.deepEqual(JSON.parse(JSON.stringify(result)), { failed: [] });
  assert.deepEqual(Array.from(model.get().announcements), [
    '本會消息一', '', '', '', '', '', '', '', '', ''
  ]);
  assert.deepEqual(Array.from(model.get().churchNews), [
    '教界消息一', '教界消息二', '', '', '', '', '', '', '', ''
  ]);
  assert.deepEqual(JSON.parse(JSON.stringify(model.get().prayer)), {
    homeRest: '',
    hospital: '住院代禱',
    other: ''
  });
});
