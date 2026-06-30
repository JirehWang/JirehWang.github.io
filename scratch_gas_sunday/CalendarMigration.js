/**
 * 教會行事曆 - 舊資料遷移
 *
 * 把舊的 3 張 sheet（聚會資料 / 事工細項 / 講道資訊）轉成新的 4 張 sheet 結構。
 *
 * 安全機制：
 *   1. 先把舊 3 張 sheet 複製成 _Backup_*_yyyyMMdd_HHmmss
 *   2. 才開始轉資料；轉失敗也不動原舊資料
 *   3. 預設「不刪除舊資料」— 你跑完 migration 確認沒問題後，再手動刪舊 sheet
 */

function cal_migrateOldData() {
  cal_setupSchema(); // 確保新 schema 就緒
  const ss = getCalendarSS();

  const oldMain     = ss.getSheetByName('聚會資料');
  const oldMinistry = ss.getSheetByName('事工細項');
  const oldSermons  = ss.getSheetByName('講道資訊');

  if (!oldMain) return { success: false, message: '❌ 找不到舊的「聚會資料」分頁，沒有東西可以遷移' };

  // 1. 備份
  const stamp = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmmss');
  const backups = [];
  [oldMain, oldMinistry, oldSermons].forEach(sh => {
    if (sh) {
      const backupName = '_Backup_' + sh.getName() + '_' + stamp;
      sh.copyTo(ss).setName(backupName);
      backups.push(backupName);
    }
  });

  // 2. 讀新結構中的「講道資訊」頂層 + 3 個子類型 + 欄位
  const types = _cal_readAll(_CAL_SHEET.TYPES);
  const sermonRoot = types.find(t => !t.parentTypeId && t['名稱'] === '講道資訊');
  if (!sermonRoot) {
    return { success: false, message: '新結構缺少「講道資訊」頂層類型，請先跑 cal_setupSchema()' };
  }
  const sermonSubs = {};
  types.filter(t => t.parentTypeId === sermonRoot.typeId).forEach(t => {
    sermonSubs[t['名稱']] = t; // 台語、華語、聯合
  });

  const fields = _cal_readAll(_CAL_SHEET.FIELDS).filter(f => f.typeId === sermonRoot.typeId);
  const fieldByName = {};
  fields.forEach(f => fieldByName[f['顯示名稱']] = f);

  // 3. 讀舊資料
  const mainData     = oldMain.getDataRange().getValues();
  const ministryData = oldMinistry ? oldMinistry.getDataRange().getValues() : [];
  const sermonData   = oldSermons ? oldSermons.getDataRange().getValues() : [];

  // 舊 主表：[ID, 日期, 聚會名稱, 聚會類別, 最後更新時間]
  // 舊 講道：[聚會ID, 講道ID, 講道類別, 講題, 講員, 經文, 宣召, 金句, 詩歌, 備註]
  const oldIdToMain = {};
  for (let i = 1; i < mainData.length; i++) {
    if (!mainData[i][0]) continue;
    oldIdToMain[mainData[i][0]] = {
      oldId:    mainData[i][0],
      date:     formatDate(mainData[i][1]),
      name:     mainData[i][2] || '',
      category: mainData[i][3] || '' // 台華語聚會 / 聯合聚會
    };
  }

  // 4. 為每筆「舊聚會」+「對應的講道」建立新事項
  const newEvents = []; // [eventId, typeId, 日期, 標題, 建立者, 建立時間, 最後更新時間]
  const newValues = []; // [eventId, fieldId, 值]
  let createdCount = 0, skippedCount = 0;
  const now = new Date();

  for (let i = 1; i < sermonData.length; i++) {
    const row = sermonData[i];
    const oldEventId = row[0];
    const sermonType = String(row[2] || '').trim(); // 台語 / 華語 / 聯合
    if (!oldEventId || !sermonType) { skippedCount++; continue; }

    const mainInfo = oldIdToMain[oldEventId];
    if (!mainInfo) { skippedCount++; continue; }

    let subType = sermonSubs[sermonType];
    if (!subType && sermonType === '聯合') {
      subType = sermonSubs['聯合-台語']; // 舊版「聯合」預設遷移至「聯合-台語」
    }
    if (!subType) { skippedCount++; continue; }

    const eventId = Utilities.getUuid();
    const title = (row[3] || '') + (row[4] ? `（${row[4]}）` : ''); // 講題（講員）
    newEvents.push([
      eventId, subType.typeId, mainInfo.date, title, 'migration', now, now
    ]);

    // 依欄位定義寫值
    const fieldValuePairs = [
      ['講題',  row[3]],
      ['講員',  row[4]],
      ['經文',  row[5]],
      ['宣召',  row[6]],
      ['金句',  row[7]],
      ['詩歌',  row[8]],
      ['備註',  row[9]]
    ];
    fieldValuePairs.forEach(([fname, fval]) => {
      const f = fieldByName[fname];
      if (f && fval !== '' && fval !== null && fval !== undefined) {
        newValues.push([eventId, f.fieldId, String(fval)]);
      }
    });
    createdCount++;
  }

  // 5. 批次寫入新表
  if (newEvents.length > 0) {
    const evSh = _cal_getSheet(_CAL_SHEET.EVENTS);
    evSh.getRange(evSh.getLastRow() + 1, 1, newEvents.length, newEvents[0].length).setValues(newEvents);
  }
  if (newValues.length > 0) {
    const vlSh = _cal_getSheet(_CAL_SHEET.VALUES);
    vlSh.getRange(vlSh.getLastRow() + 1, 1, newValues.length, newValues[0].length).setValues(newValues);
  }

  return {
    success: true,
    message: `✅ 遷移完成：${createdCount} 個事項已建立，${skippedCount} 筆跳過`,
    backups: backups,
    note: '舊「聚會資料 / 事工細項 / 講道資訊」尚未刪除，請手動確認沒問題後再刪除。'
  };
}

// 清掉新結構所有資料（事項 + 欄位值），保留類型與欄位定義（測試用）
function cal_clearNewData() {
  [_CAL_SHEET.EVENTS, _CAL_SHEET.VALUES].forEach(name => {
    const sh = _cal_getSheet(name);
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
  });
  return { success: true, message: '已清空所有事項與欄位值（類型與欄位定義保留）' };
}
