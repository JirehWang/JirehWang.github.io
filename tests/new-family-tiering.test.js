const test = require('node:test');
const assert = require('node:assert/strict');

const mod = require('../apps/LKC_NewFamily/new-family-supabase.js');
const {
  NewFamilySupabaseService,
  HOT_YEAR_THRESHOLD,
  isHotYear,
  isHotDate,
  isColdCase
} = mod;

test('新家人冷熱分流常數設定為 2025', () => {
  assert.equal(HOT_YEAR_THRESHOLD, 2025, 'HOT_YEAR_THRESHOLD 應為 2025');
  assert.equal(isHotYear(2024), false, '2024 年應為冷資料');
  assert.equal(isHotYear(2025), true, '2025 年應為熱資料');
  assert.equal(isHotYear(2026), true, '2026 年應為熱資料');

  assert.equal(isHotDate('2024-10-06'), false, '2024-10-06 應判定為冷資料');
  assert.equal(isHotDate('2025-01-01'), true, '2025-01-01 應判定為熱資料');
  assert.equal(isHotDate('2026-04-05'), true, '2026-04-05 應判定為熱資料');
});

test('isColdCase 正確識別 2024 年歷史個案', () => {
  const case2024 = {
    '表單號': '20241006201',
    '姓名': '楊小貞',
    '首次來訪日': '2024-10-06'
  };
  const case2025 = {
    '表單號': '20250105001',
    '姓名': '張小明',
    '首次來訪日': '2025-01-05'
  };
  const case2026 = {
    '表單號': '20260405008',
    '姓名': '葉大德',
    '首次來訪日': '2026-04-05'
  };

  assert.equal(isColdCase(case2024), true, '2024 個案應判定為冷資料');
  assert.equal(isColdCase(case2025), false, '2025 個案不應判定為冷資料');
  assert.equal(isColdCase(case2026), false, '2026 個案不應判定為冷資料');
});

test('getClosedCases 純冷區間查詢透明轉發至冷端', async () => {
  let calledCold = false;
  global.window = {
    churchAPI_original_nf: async (action, payload) => {
      if (action === 'getClosedCases') {
        calledCold = true;
        return {
          success: true,
          data: [{ '表單號': '20241006201', '姓名': '楊小貞', '首次來訪日': '2024-10-06' }]
        };
      }
    }
  };

  const res = await NewFamilySupabaseService.getClosedCases({
    startDate: '2024-01-01',
    endDate: '2024-12-31'
  });

  assert.equal(calledCold, true, '純 2024 查詢應轉發至冷端');
  assert.equal(res.data.length, 1);
  assert.equal(res.data[0]['姓名'], '楊小貞');
});

test('getClosedCases 全量查詢透明合併冷熱資料並去重排序', async () => {
  // Mock Supabase client
  const mockHotData = [
    { id: '1', form_number: '20260405001', name: '王熱門', first_visit_date: '2026-04-05', status: 'closed' },
    { id: '2', form_number: '20250501001', name: '李熱門', first_visit_date: '2025-05-01', status: 'closed' }
  ];

  global.window = {
    _supabase: {
      from: () => ({
        select: () => ({
          eq: () => ({
            gte: () => ({
              order: async () => ({ data: mockHotData, error: null })
            })
          })
        })
      })
    },
    churchAPI_original_nf: async (action, payload) => {
      return {
        success: true,
        data: [
          { '表單號': '20241006201', '姓名': '楊小貞', '首次來訪日': '2024-10-06' },
          { '表單號': '20250501001', '姓名': '李熱門舊版', '首次來訪日': '2025-05-01' } // should be deduplicated by hot
        ]
      };
    }
  };

  const res = await NewFamilySupabaseService.getClosedCases({});
  assert.equal(res.success, true);
  assert.equal(res.data.length, 3, '應合併 2 筆熱資料 + 1 筆冷資料 (2025 重複項以熱資料優先)');
  assert.equal(String(res.data[0]['表單號']), '20260405001', '第一筆應為 2026 最新個案');
  assert.equal(String(res.data[1]['表單號']), '20250501001');
  assert.equal(res.data[1]['姓名'], '李熱門', '重複表單號應保留熱資料版本');
  assert.equal(String(res.data[2]['表單號']), '20241006201', '第三筆應為 2024 歷史個案');
});

