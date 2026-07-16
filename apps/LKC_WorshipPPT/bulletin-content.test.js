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

test('keeps announcements and pastoral prayer as separate report page types', () => {
  const pages = buildReportPages({
    announcements: ['消息一', '消息二', '消息三', '消息四'],
    prayer: { homeRest: '王小明', hospital: '陳小華', other: '為社區代禱' }
  });

  assert.equal(pages.length, 3);
  assert.equal(pages[0].kind, 'report');
  assert.equal(pages[0].title, '報告－本會消息');
  assert.match(pages[0].body, /1\. 消息一/);
  assert.match(pages[1].body, /4\. 消息四/);
  assert.equal(pages[2].title, '報告－關懷代禱');
  assert.match(pages[2].body, /在家調養兄姐：王小明/);
  assert.match(pages[2].body, /住院：陳小華/);
  assert.match(pages[2].body, /其他代禱：為社區代禱/);
});

test('applies report pages and praise fields without treating cloud values as a single report string', () => {
  const model = {
    announcements: { title: '報告', body: '' },
    praise: { title: '讚美', kicker: '', body: '' }
  };

  applyReportsToModel(model, {
    announcements: ['消息一', '', '消息二'],
    prayer: { homeRest: '在家調養名單', hospital: '', other: '' }
  });
  applyPraiseToModel(model, {
    title: '讚美無停',
    kicker: '聖歌隊',
    lyrics: '第一段\n\n第二段'
  });

  assert.deepEqual(model.announcements.announcements, ['消息一', '消息二']);
  assert.equal(model.announcements.includeSectionTitle, true);
  assert.equal(model.announcements.pptPages.length, 2);
  assert.equal(model.praise.title, '讚美無停');
  assert.equal(model.praise.kicker, '聖歌隊');
  assert.equal(model.praise.body, '第一段\n\n第二段');
});
