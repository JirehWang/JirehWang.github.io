const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const html = fs.readFileSync(path.join(__dirname, '..', 'admin.html'), 'utf8');

const expectedCards = [
  ['apps/LKC_MinistrySchedule/', '教會事工總表系統'],
  ['apps/LKC_worship/admin.html', '敬拜團服事表'],
  ['apps/LKC_ppt_generator/', '聖經PPT產生器'],
  ['apps/LKC_PrayerPPT/', '禱告會PPT產生器'],
  ['apps/LKC_TaiwaneseAudioBible/', '台語有聲聖經'],
  ['apps/LKC_WhosCar/', '教會車牌管理系統'],
  ['apps/LKC_Group/', '教會小組點名系統'],
  ['apps/LKC_SundayserviceAttendance/', '教會主日出席點名系統'],
  ['apps/LKC_ChildrenAttendance/', '兒童主日出席點名系統'],
  ['https://lkcweekly.netlify.app/#section-9B23KUaK0Fu91NerT8/dev', '教會電子週報'],
  ['apps/LKC_MasterSchedule/', '教會行事曆管理系統'],
  ['apps/LKC_SundayBulletin/', '教會週報管理系統'],
  ['apps/LKC_Offering/', '奉獻管理系統'],
  ['apps/LKC_NewFamily/', '新家人管理系統'],
  ['apps/LKC_MemberStatus/', '同工服事狀態管理系統']
];

test('admin portal preserves all existing cards, destinations, and new-tab behavior', () => {
  for (const [href, label] of expectedCards) {
    assert.match(html, new RegExp(`<a class="card" href="${href.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}" target="_blank">`));
    assert.ok(html.includes(`<div class="title">${label}</div>`), label);
  }
});

test('admin portal preserves cache action labels and confirmations', () => {
  for (const label of [
    '小組點名資料', '主日出席資料', '事工管理資料', '敬拜團資料',
    '教會行事曆資料', '新家人資料', '同工服事狀態管理資料',
    '全部資料快取', '修復前端版本 / PWA'
  ]) {
    assert.ok(html.includes(label), label);
  }
  assert.ok(html.includes('這會清除所有系統的資料快取'));
  assert.ok(html.includes('這會清除本機 PWA/前端快取並重新載入'));
});

test('admin portal loads the cache coordinator before invoking it', () => {
  const scriptIndex = html.indexOf('src="admin-cache-coordinator.js"');
  const usageIndex = html.indexOf('window.AdminCacheCoordinator.refreshCacheGroup');
  assert.ok(scriptIndex >= 0);
  assert.ok(usageIndex > scriptIndex);
});
