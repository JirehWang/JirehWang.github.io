const assert = require('assert');

function makeSheet(values) {
  return {
    name: '',
    getName() {
      return this.name;
    },
    setName(name) {
      this.name = name;
      return this;
    },
    getLastRow() {
      return values.length;
    },
    getLastColumn() {
      return values[0] ? values[0].length : 0;
    },
    getDataRange() {
      return {
        getValues() {
          return values;
        }
      };
    }
  };
}

function namedSheet(name, values) {
  return makeSheet(values).setName(name);
}

const ministrySheets = {
  Config: makeSheet([
    ['UUID', 'ID', '名稱', '模板', '狀態', '規則', '名單', '講道設定', 'pageFieldConfig'],
    ['uuid-a', 'G01', '葡萄樹A組', '小組聚會表模板', '啟用', '', '[]', '', ''],
    ['uuid-b', 'M01', '招待組', '事工型模板', '啟用', '', JSON.stringify([{ name: '王小明' }, { name: '重名' }]), '', JSON.stringify({ scheduleMode: 'schedule' })],
    ['uuid-d', 'M02', '關懷組', '事工型模板', '啟用', '', JSON.stringify([{ name: '王小明' }]), '', JSON.stringify({ scheduleMode: 'membersOnly' })],
    ['uuid-c', 'F01', '團契A', '團契聚會表模板', '啟用', '', '[]', '', '']
  ]),
  '葡萄樹A組': makeSheet([
    ['日期', '主題', '破冰', '敬拜', '備註'],
    ['2026-06-01', '聚會一', '王小明', '李小華', ''],
    ['2024-12-01', '太舊', '王小明', '', '']
  ]),
  '招待組': makeSheet([
    ['姓名', '備註', '日期', '班表欄位'],
    ['王小明', '', '2026-06-01', '王小明'],
    ['不存在', '', '2026-06-01', '不應讀取']
  ]),
  '關懷組': makeSheet([
    ['姓名', '備註', '日期', '班表欄位'],
    ['王小明', '', '2026-06-01', '不應讀取']
  ]),
  '團契A': makeSheet([
    ['日期', '司會', '敬拜'],
    ['2026-03-15', '陳美麗', '重名']
  ])
};

global.CacheService = {
  getScriptCache() {
    return {
      get() { return null; },
      put() {},
      remove() {}
    };
  }
};
global.getCachedMembers = () => [
  ['王小明', '男', '', '備註', '', '2026-01-01', '', 'LK00001', '葡萄樹A組', '核心同工'],
  ['李小華', '女', '', '', '', '2026-01-01', '', 'LK00002', '葡萄樹A組', '一般同工'],
  ['陳美麗', '女', '', '', '', '2026-01-01', '', 'LK00003', '團契A', '核心同工'],
  ['不統計人', '男', '', '', 'TRUE', '2026-01-01', '', 'LK00999', '葡萄樹A組', '小羊'],
  ['重名', '男', '', '', '', '2026-01-01', '', 'LK00004', '', '小羊'],
  ['重名', '女', '', '', '', '2026-01-01', '', 'LK00005', '', '小羊']
];
global.parseGroupString = value => String(value || '').split(/[、,，]/).map(s => s.trim()).filter(Boolean);
global.parseGroupRoles = (groupStr, roleStr) => {
  const groups = global.parseGroupString(groupStr);
  const result = {};
  groups.forEach(g => { result[g] = roleStr || '小羊'; });
  return result;
};
global.ensureConfigSchemaV3 = () => {};
global._getConfigData = () => ministrySheets.Config.getDataRange().getValues();
global.getMinistrySS = () => ({
  getSheetByName(name) {
    return ministrySheets[name] || null;
  }
});
const attendanceSheets = {
  '台語點名紀錄': namedSheet('台語點名紀錄', [
    ['出席日', '名單', '新朋友(男)', '新朋友(女)'],
    ['2026-06-07', 'LK00001, LK00002, LK00999', 0, 0],
    ['2024-05-01', 'LK00001', 0, 0]
  ]),
  '主日學A班點名紀錄': namedSheet('主日學A班點名紀錄', [
    ['出席日', '名單', '新朋友(男)', '新朋友(女)'],
    ['2026-06-07', 'LK00002', 0, 0],
    ['2026-06-14', 'LK00002', 0, 0]
  ])
};
global.getSS = () => ({
  getSheetByName(name) {
    return attendanceSheets[name] || null;
  },
  getSheets() {
    return Object.keys(attendanceSheets).map(name => attendanceSheets[name]);
  }
});
const groupAttendanceSheets = {
  '葡萄樹A組_點名紀錄': namedSheet('葡萄樹A組_點名紀錄', [
    ['日期', '出席', '缺席', '新朋友', '總數'],
    ['2026-06-08', 'LK00001, LK00999', 'LK00002', '', 1],
    ['2024-06-08', 'LK00001', '', '', 1]
  ]),
  '團契A_點名紀錄': namedSheet('團契A_點名紀錄', [
    ['日期', '出席', '缺席', '新朋友', '總數'],
    ['2026-06-08', 'LK00003', '', '', 1]
  ])
};
global.getGroupSS = () => ({
  getSheets() {
    return Object.keys(groupAttendanceSheets).map(name => groupAttendanceSheets[name]);
  }
});
global.getScheduleByDateRange = () => ({
  status: 'success',
  data: [
    { '日期': '2026-05-01', '聚會名稱': '主日', '聚會類別': '華語', '主領': '王小明', '司琴': '李小華' },
    { '日期': '2024-01-01', '聚會名稱': '舊資料', '聚會類別': '華語', '主領': '王小明' },
    { '日期': '2026-05-08', '聚會名稱': '主日', '聚會類別': '華語', '主領': '不存在' }
  ]
});

const core = require('../../../LKC/主日出席_測試版/MemberStatusCore.js');
const aggregate = core._memberStatusBuildAggregate({ now: '2026-07-02T12:00:00+08:00' });
assert.strictEqual(aggregate.members.some(m => m.uid === 'LK00999'), false, 'excluded members should not appear in member status aggregate');

const wang = aggregate.members.find(m => m.uid === 'LK00001');
assert(wang, '王小明 profile should exist');
assert.strictEqual(wang.groups[0].name, '葡萄樹A組');
assert.strictEqual(wang.groups[0].role, '核心同工');
assert.strictEqual(wang.groupMinistries.length, 1, 'group ministry should be attached');
assert.strictEqual(wang.groupMinistries[0].duties.includes('破冰'), true, 'recent group duty should be read');
assert.strictEqual(wang.groupMinistries[0].serviceHistory.length, 1, 'old group rows should be excluded');
assert.strictEqual(wang.churchMinistries.length, 2, 'non-group ministry membership should be attached');
assert.strictEqual(wang.churchMinistries[0].ministryName, '招待組');
assert.strictEqual(wang.churchMinistries[0].serviceHistory.length, 1, 'scheduled non-group ministry schedules should be read');
assert.strictEqual(wang.churchMinistries[0].duties.includes('班表欄位'), true, 'scheduled non-group ministry duties should be tracked');
const careMinistry = wang.churchMinistries.find(x => x.ministryName === '關懷組');
assert(careMinistry, 'members-only ministry membership should be attached');
assert.strictEqual(careMinistry.serviceHistory.length, 0, 'members-only ministry schedules must not be read');
assert.strictEqual(wang.worship.positions.includes('主領'), true, 'worship schedule should be read');
assert.strictEqual(wang.worship.serviceHistory.length, 1, 'old worship rows should be excluded');
assert.strictEqual(wang.discipleship.status, 'unknown');
assert.strictEqual(wang.attendance.sunday.count, 1, 'recent Sunday service attendance should be counted');
assert.strictEqual(wang.attendance.sunday.total, 1, 'old Sunday service sessions should be excluded');
assert.strictEqual(wang.attendance.sundaySchool.count, 0, 'Sunday school absence should be counted separately');
assert.strictEqual(wang.attendance.group.count, 1, 'recent assigned group attendance should be counted');
assert.strictEqual(wang.attendance.group.total, 1, 'old group attendance sessions should be excluded');

const zhang = aggregate.members.find(m => m.uid === 'LK00002');
assert.strictEqual(zhang.attendance.sunday.count, 1, 'Sunday attendance should attach to each present UID');
assert.strictEqual(zhang.attendance.sundaySchool.count, 2, 'Sunday school attendance should attach independently');
assert.strictEqual(zhang.attendance.group.count, 0, 'group absence should be reflected');

const wangSummary = core._memberStatusSummarizeProfile(wang);
assert.strictEqual(wangSummary.attendance.sunday.count, 1, 'summary should expose attendance');
assert.strictEqual(wangSummary.participation.ministryCount >= 3, true, 'summary should expose ministry participation count');

const sundaySchoolOnly = core._memberStatusApplyMemberFilters(aggregate.members, {
  attendanceSunday: 'present',
  attendanceSundaySchool: 'present',
  attendanceGroup: 'absent'
});
assert.deepStrictEqual(sundaySchoolOnly.map(m => m.uid), ['LK00002'], 'attendance filters should intersect across sources');

const unresolvedDuplicate = aggregate.unresolvedParticipants.find(p => p.name === '重名' && p.reason === 'duplicate-name');
assert(unresolvedDuplicate, 'duplicate names should be unresolved');
const unresolvedMissing = aggregate.unresolvedParticipants.find(p => p.name === '不存在' && p.reason === 'not-found');
assert(unresolvedMissing, 'missing names should be unresolved');

console.log('member-status-core tests passed');
