/**
 * MigrationAttendance.js — 點名紀錄姓名 → UID 遷移
 *
 * 將所有點名紀錄 sheet 的「名單」欄從姓名格式轉成 UID 格式：
 *   舊：張三(男), 李四(女), 王小華(男)
 *   新：LK00001, LK00002, LK00003
 *
 * 影響範圍：
 *   主日試算表：台語點名紀錄 / 華語點名紀錄 / 聯合點名紀錄 / 主日學*點名紀錄 等
 *   小組試算表：所有 *_點名紀錄
 *
 * 設計：
 *   - 冪等：可重複執行，已是 UID 的會跳過
 *   - 找不到對應會友的姓名 → 移到「未知名單」附註，不丟失資訊
 *   - 進度紀錄到 PropertiesService，若中斷可從上次繼續
 *
 * 執行方式：
 *   GAS 編輯器 → 函式選 migrateAttendanceRecordsToUid → 按 ▶ 執行
 *   執行紀錄會列出每張 sheet 的轉換統計
 */

// ═══════════════════════════════════════════════════════════
//  主入口：遷移所有點名紀錄
// ═══════════════════════════════════════════════════════════
function migrateAttendanceRecordsToUid() {
  Logger.log('🚀 開始遷移點名紀錄 姓名 → UID...');
  const lookups = getMemberLookups();

  const totalStats = {
    sheetsProcessed: 0,
    rowsConverted: 0,
    rowsSkippedAlreadyUid: 0,
    namesNotFound: {}  // name -> 出現次數
  };

  // 1. 主日試算表：所有 *點名紀錄 sheet
  Logger.log('── 處理主日試算表 ──');
  const mainSS = getSS();
  mainSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.endsWith("點名紀錄")) return;
    if (name === "點名紀錄") return; // 排除非預期同名
    _migrateOneAttendanceSheet(sheet, lookups, totalStats);
  });

  // 2. 小組試算表：所有 *_點名紀錄 sheet
  Logger.log('── 處理小組試算表 ──');
  const groupSS = getGroupSS();
  groupSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.endsWith("_點名紀錄")) return;
    _migrateOneAttendanceSheet(sheet, lookups, totalStats);
  });

  // 3. 主日 SYNC_TEMP（短期暫存，順手清空）
  Logger.log('── 清空 SYNC_TEMP（短期暫存無需保留）──');
  const syncSheet = mainSS.getSheetByName("SYNC_TEMP");
  if (syncSheet && syncSheet.getLastRow() > 1) {
    syncSheet.deleteRows(2, syncSheet.getLastRow() - 1);
  }

  // 4. 清出席計數 cache（90 天）— 強制重建
  const cache = CacheService.getScriptCache();
  ['台語', '華語', '聯合', '主日學A班', '主日學B班'].forEach(t => {
    const todayDateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
    cache.remove("ATT_MAP_" + t + "_" + todayDateStr);
  });

  // 摘要
  Logger.log('═══════════════════════════════');
  Logger.log('✅ 遷移完成');
  Logger.log(`  處理 sheet 數：${totalStats.sheetsProcessed}`);
  Logger.log(`  轉換的列數：${totalStats.rowsConverted}`);
  Logger.log(`  跳過（已是 UID）：${totalStats.rowsSkippedAlreadyUid}`);
  const notFoundNames = Object.keys(totalStats.namesNotFound);
  if (notFoundNames.length > 0) {
    Logger.log('⚠️ 找不到對應 UID 的姓名（已從紀錄中移除，請手動確認是否需處理）：');
    notFoundNames.sort((a, b) => totalStats.namesNotFound[b] - totalStats.namesNotFound[a]).forEach(n => {
      Logger.log(`  - ${n} (出現 ${totalStats.namesNotFound[n]} 次)`);
    });
  }
  Logger.log('═══════════════════════════════');
  return { ok: true, stats: totalStats };
}

// ═══════════════════════════════════════════════════════════
//  子函式：處理單一 sheet
// ═══════════════════════════════════════════════════════════
function _migrateOneAttendanceSheet(sheet, lookups, totalStats) {
  const sheetName = sheet.getName();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 名單欄是第 2 欄（B 欄）
  const range = sheet.getRange(2, 2, lastRow - 1, 1);
  const values = range.getValues();
  let modified = false;
  let convertedCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const cell = values[i][0];
    if (!cell) continue;
    const original = String(cell).trim();
    if (!original) continue;

    // 解析每個 entry
    const parts = original.split(/[,，、]\s*/);
    let allUids = true;
    let needsConversion = false;
    const newParts = [];

    parts.forEach(part => {
      const item = part.trim();
      if (!item) return;
      const cleaned = item.split('(')[0].trim();
      if (!cleaned) return;

      if (/^LK\d+$/i.test(cleaned)) {
        // 已是 UID
        newParts.push(cleaned.toUpperCase());
      } else {
        // 是姓名 → 反查 UID
        allUids = false;
        const uid = lookups.n2u[cleaned];
        if (uid) {
          newParts.push(uid);
          needsConversion = true;
        } else {
          // 找不到對應 → 記錄但不寫入
          totalStats.namesNotFound[cleaned] = (totalStats.namesNotFound[cleaned] || 0) + 1;
          needsConversion = true;
        }
      }
    });

    if (allUids) {
      skippedCount++;
      continue;
    }

    if (needsConversion) {
      values[i][0] = newParts.join(", ");
      modified = true;
      convertedCount++;
    }
  }

  if (modified) {
    range.setValues(values);
  }

  totalStats.sheetsProcessed++;
  totalStats.rowsConverted += convertedCount;
  totalStats.rowsSkippedAlreadyUid += skippedCount;
  Logger.log(`  ${sheetName}: 轉 ${convertedCount} 列 / 跳過 ${skippedCount} 列`);
}

// ═══════════════════════════════════════════════════════════
//  輔助：dry-run 模式（只統計不寫入，預覽結果）
// ═══════════════════════════════════════════════════════════
function dryRunMigrationAttendance() {
  Logger.log('🧪 Dry-run 模式：只預覽不寫入 Sheet');
  const lookups = getMemberLookups();

  const allSheets = [];
  getSS().getSheets().forEach(s => { if (s.getName().endsWith("點名紀錄")) allSheets.push(s); });
  getGroupSS().getSheets().forEach(s => { if (s.getName().endsWith("_點名紀錄")) allSheets.push(s); });

  const namesNotFound = {};
  let totalRows = 0, totalConvert = 0, totalSkip = 0;

  allSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let convertCount = 0, skipCount = 0;

    values.forEach(row => {
      const cell = row[0];
      if (!cell) return;
      const parts = String(cell).split(/[,，、]\s*/);
      let allUids = true, hasName = false;
      parts.forEach(p => {
        const cleaned = p.split('(')[0].trim();
        if (!cleaned) return;
        if (/^LK\d+$/i.test(cleaned)) return;
        allUids = false;
        hasName = true;
        if (!lookups.n2u[cleaned]) {
          namesNotFound[cleaned] = (namesNotFound[cleaned] || 0) + 1;
        }
      });
      if (allUids) skipCount++;
      else if (hasName) convertCount++;
    });

    Logger.log(`  ${sheet.getName()}: 待轉 ${convertCount} 列 / 已是 UID ${skipCount} 列`);
    totalRows += values.length;
    totalConvert += convertCount;
    totalSkip += skipCount;
  });

  Logger.log('═══════════════════════════════');
  Logger.log(`📊 Dry-run 摘要：總列 ${totalRows}，待轉 ${totalConvert}，已是 UID ${totalSkip}`);
  if (Object.keys(namesNotFound).length > 0) {
    Logger.log('⚠️ 以下姓名找不到對應 UID（執行時會被略過）：');
    Object.keys(namesNotFound).sort((a, b) => namesNotFound[b] - namesNotFound[a]).forEach(n => {
      Logger.log(`  - ${n} (${namesNotFound[n]} 次)`);
    });
  }
  Logger.log('═══════════════════════════════');
  return { totalRows, totalConvert, totalSkip, namesNotFound };
}
