const test = require('node:test');
const assert = require('node:assert/strict');

const {
  buildBulletinCloudUrl,
  buildReportPages,
  applyReportsToModel,
  applyPraiseToModel
} = require('./bulletin-content.js');

test('uses the Sunday bulletin cloud keys for the selected service date', () => {
  assert.match(buildBulletinCloudUrl('https://example.test/exec', 'reports', '2026-07-12'), /key=reports_2026-07-12/);
  assert.match(buildBulletinCloudUrl('https://example.test/exec', 'praise', '2026-07-12'), /key=praise_songs_2026-07-12/);
});

test('keeps announcements, church news, and pastoral prayer in report-page order', () => {
  const pages = buildReportPages({
    announcements: ['消息一', '消息二', '消息三', '消息四'],
    churchNews: ['教界一', '教界二', '教界三'],
    prayer: { homeRest: '王小明', hospital: '陳小華', other: '為社區代禱' }
  });

  assert.equal(pages.length, 4);
  assert.equal(pages[0].kind, 'report');
  assert.equal(pages[0].title, '報告－本會消息');
  assert.match(pages[0].body, /1\. 消息一/);
  assert.match(pages[1].body, /4\. 消息四/);
  assert.equal(pages[2].title, '報告－教界消息');
  assert.match(pages[2].body, /1\. 教界一/);
  assert.match(pages[2].body, /3\. 教界三/);
  assert.equal(pages[3].title, '報告－關懷代禱');
  assert.match(pages[3].body, /在家調養兄姐：王小明/);
  assert.match(pages[3].body, /住院：陳小華/);
  assert.match(pages[3].body, /其他代禱：為社區代禱/);
});

test('paginates the 2026-06-28 announcements by rendered line capacity', () => {
  const pages = buildReportPages({
    announcements: [
      '謝謝李俊佑牧師的信息分享。',
      '下午一點召開定期小會，請相關同工預備心參加。',
      '教會第二屆雙翼門徒訓練畢業名單：出如玉、謝育倫、周倩如、蔡君宜、曾宇璿、劉秀琴、甘淑蘭、林恩予、杜文心、林慧敏、金仕淳、徐聆、陳美如、羅彩珍，共14位。下主日禮拜中舉行畢業典禮，當天為台華語聯合禮拜及聯合成人主日學，會後備有愛餐(芥菜種)，請兄姊自備餐具。',
      '下主日下午一點召開定期長執會請同工預備心參加。',
      '有訂購每日讀經釋義的兄姊，請至辦公室領取第三季的讀本。'
    ],
    churchNews: [],
    prayer: {}
  });

  assert.equal(pages.length, 4);
  assert.match(pages[0].body, /1\. 謝謝/);
  assert.match(pages[0].body, /2\. 下午一點/);
  assert.doesNotMatch(pages[0].body, /3\./);
  assert.match(pages[1].body, /^3\. /);
  assert.doesNotMatch(pages[1].body, /4\./);
  assert.match(pages[2].body, /^3\.（續）/);
  assert.doesNotMatch(pages[2].body, /4\./);
  assert.match(pages[3].body, /4\. 下主日/);
  assert.match(pages[3].body, /5\. 有訂購/);
  pages.forEach(page => assert.ok(page.body.split('\n').length <= 5));
});

test('applies report pages and praise fields without treating cloud values as a single report string', () => {
  const model = {
    announcements: { title: '報告', body: '' },
    praise: { title: '讚美', kicker: '', body: '' }
  };

  applyReportsToModel(model, {
    announcements: ['消息一', '', '消息二'],
    churchNews: ['教界一', '', '教界二'],
    prayer: { homeRest: '在家調養名單', hospital: '', other: '' }
  });
  applyPraiseToModel(model, {
    title: '讚美無停',
    kicker: '聖歌隊',
    lyrics: '第一段\n\n第二段'
  });

  assert.deepEqual(model.announcements.announcements, ['消息一', '消息二']);
  assert.deepEqual(model.announcements.churchNews, ['教界一', '教界二']);
  assert.equal(model.announcements.includeSectionTitle, true);
  assert.equal(model.announcements.pptPages.length, 3);
  assert.equal(model.praise.title, '讚美無停');
  assert.equal(model.praise.kicker, '聖歌隊');
  assert.equal(model.praise.body, '第一段\n\n第二段');
});
