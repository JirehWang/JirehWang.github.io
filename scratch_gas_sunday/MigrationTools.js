/**
 * MigrationTools.js — 一次性資料遷移工具
 *
 * 整合方案 C 上線前需手動執行一次的函式（在 GAS 編輯器選擇後按 ▶）
 * 全部設計為「冪等」：可以重複執行，不會搞壞資料
 */

// ═══════════════════════════════════════════════════════════
//  migrateGroupRolesToMaster
//
//  把所有小組 _名單 sheet 的「身分」資料遷移到主日 會友名單，
//  支援「一個會友屬多組、每組可有不同身分」：
//
//    所屬小組欄：用「、」分隔     例：葡萄樹A組、以斯帖小組
//    身分欄：    用「、」分隔，含括號標註所屬組
//                 單組：核心同工
//                 多組：核心同工(葡萄樹A組)、一般同工(以斯帖小組)
//
//  邏輯：
//    1. 掃所有 *_名單 sheet → 收集 { 姓名: { 小組: 身分 } } 結構
//    2. 同名同組多次出現：以「優先級高」為準（核心>一般>陪伴>小羊）
//    3. 跟主日現有資料合併：
//         - 已存在 → 解析現有 group/role 字串 → 合併新資料 → 寫回
//         - 不存在 → 自動加入主日，含全部小組與身分
//    4. 重建快取
//
//  執行方式：
//    GAS 編輯器 → 函式選 migrateGroupRolesToMaster → 按 ▶ 執行
// ═══════════════════════════════════════════════════════════
function migrateGroupRolesToMaster() {
  Logger.log('🚀 開始遷移小組身分資料到主日名單...');

  const masterSheet = getMemberSheet();
  const masterData  = masterSheet.getDataRange().getValues();

  // 建立 主日 名單索引：姓名 -> { rowIndex, groupRoles, raw }
  const masterIndex = {};
  for (let i = 1; i < masterData.length; i++) {
    const name = masterData[i][0] ? String(masterData[i][0]).trim() : "";
    if (!name) continue;
    masterIndex[name] = {
      rowIndex: i + 1,
      groupRoles: parseGroupRoles(masterData[i][9], masterData[i][10])  // helpers in MemberDB.js
    };
  }
  Logger.log('📖 主日名單已索引：' + Object.keys(masterIndex).length + ' 位會友');

  // ── 第一階段：收集所有小組 _名單 的資料 ──
  // collected = { 姓名: { 小組名: 身分 } }
  const collected = {};
  const PRI = { "核心同工": 4, "一般同工": 3, "陪伴同工": 2, "小羊": 1 };

  const groupSS = getGroupSS();
  groupSS.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (!sheetName.endsWith("_名單")) return;
    const groupName = sheetName.replace("_名單", "");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    rows.forEach(row => {
      const name = row[0] ? String(row[0]).trim() : "";
      const roleRaw = row[2] ? String(row[2]).trim() : "小羊";
      if (!name) return;
      const role = (MEMBER_ROLES.indexOf(roleRaw) !== -1) ? roleRaw : "小羊";

      if (!collected[name]) collected[name] = {};
      // 同人同組多次：取優先級高的
      if (collected[name][groupName]) {
        if ((PRI[role] || 1) > (PRI[collected[name][groupName]] || 1)) {
          collected[name][groupName] = role;
        }
      } else {
        collected[name][groupName] = role;
      }
    });
  });

  Logger.log(`📦 從小組 _名單 收集到 ${Object.keys(collected).length} 位會友`);

  // ── 第二階段：與主日合併 ──
  const stats = { added: 0, updated: 0, unchanged: 0 };
  const conflicts = [];
  const newRows = [];

  // 計算下一個 LK 編號起點
  let nextUidNum = 0;
  for (let i = 1; i < masterData.length; i++) {
    const uid = masterData[i][7] ? String(masterData[i][7]).trim() : "";
    if (uid.startsWith("LK")) {
      const num = parseInt(uid.replace("LK", ""), 10);
      if (!isNaN(num) && num > nextUidNum) nextUidNum = num;
    }
  }

  Object.keys(collected).forEach(name => {
    const newGroupRoles = collected[name];
    const master = masterIndex[name];

    if (master) {
      // ─── 已在主日：合併 groupRoles ───
      const merged = Object.assign({}, master.groupRoles);
      let changed = false;

      Object.keys(newGroupRoles).forEach(g => {
        const existingRole = merged[g];
        const newRole = newGroupRoles[g];
        if (!existingRole) {
          merged[g] = newRole;
          changed = true;
        } else if (existingRole !== newRole) {
          // 衝突：以優先級高為準
          const final = (PRI[newRole] || 1) > (PRI[existingRole] || 1) ? newRole : existingRole;
          if (final !== existingRole) {
            merged[g] = final;
            changed = true;
            conflicts.push(`${name} @ ${g}: ${existingRole} → ${final}（小組_名單為 ${newRole}）`);
          }
        }
      });

      if (changed) {
        const formatted = formatGroupRoles(merged);
        masterSheet.getRange(master.rowIndex, 10).setValue(formatted.groupStr);
        masterSheet.getRange(master.rowIndex, 11).setValue(formatted.roleStr);
        master.groupRoles = merged;
        stats.updated++;
      } else {
        stats.unchanged++;
      }
    } else {
      // ─── 主日沒有：加入新列 ───
      nextUidNum++;
      const uid = "LK" + nextUidNum.toString().padStart(5, '0');
      const now = new Date();
      const formatted = formatGroupRoles(newGroupRoles);
      // [姓名, 性別, 建立日期, 備註, 不列入統計, 異動日期, 異動紀錄, 系統編號, QR Code, 所屬小組, 身分]
      newRows.push([name, "", now, "由小組_名單自動匯入", false, now, "migrate", uid, "", formatted.groupStr, formatted.roleStr]);
      stats.added++;
    }
  });

  // 批次寫入新增會友
  if (newRows.length > 0) {
    const startRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(startRow, 1, newRows.length, 11).setValues(newRows);
  }

  invalidateAndRebuildMemberCache();

  // 摘要報告
  Logger.log('═══════════════════════════════');
  Logger.log('✅ 遷移完成');
  Logger.log(`  新增會友到主日名單：${stats.added} 位`);
  Logger.log(`  更新既有會友：${stats.updated} 位`);
  Logger.log(`  無異動：${stats.unchanged} 位`);
  if (conflicts.length > 0) {
    Logger.log('⚠️ 衝突紀錄（已用優先級規則處理）：');
    conflicts.forEach(c => Logger.log('  - ' + c));
  }
  Logger.log('═══════════════════════════════');
  return { ok: true, stats: stats, conflicts: conflicts };
}

// ═══════════════════════════════════════════════════════════
//  cleanupOldMemberCacheTrigger
//
//  方案 C 啟用後，原本獨立的 syncMemberCacheFromSheet trigger
//  已由 keepWarm 兜底取代。執行一次清掉舊 trigger 即可省 quota。
// ═══════════════════════════════════════════════════════════
function cleanupOldMemberCacheTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncMemberCacheFromSheet') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(`✅ 已清除 ${removed} 個 syncMemberCacheFromSheet trigger`);
  return removed;
}

// ═══════════════════════════════════════════════════════════
//  findDuplicateMembers
//
//  掃描主日會友名單，找出：
//   ① 完全同名（多位會友姓名一字不差）
//   ② 高度相似（Levenshtein 距離 = 1，例如「王小明」vs「王小銘」）
//   ③ 共同前綴 ≥ 2 字的（同姓同名首字，可能是親屬）— 僅資訊用
//
//  執行：GAS 編輯器選 findDuplicateMembers → ▶ 執行
//  Logger 會印出：姓名 + 系統編號 + 性別 + 所屬小組
//  你可以直接複製訊息來判斷哪些要合併、刪除或保留
// ═══════════════════════════════════════════════════════════
function findDuplicateMembers() {
  Logger.log('🔍 開始掃描會友名單重名與相似名...');

  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('⚠️ 會友名單為空');
    return;
  }

  // 索引：name -> [{ uid, gender, group, role, rowIndex }]
  const byName = {};
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0] ? String(data[i][0]).trim() : "";
    if (!name) continue;
    if (!byName[name]) byName[name] = [];
    byName[name].push({
      uid:    data[i][7]  ? String(data[i][7]).trim() : "(無編號)",
      gender: data[i][1]  ? String(data[i][1]).trim() : "?",
      group:  data[i][9]  ? String(data[i][9]).trim() : "(無組)",
      role:   data[i][10] ? String(data[i][10]).trim() : "(未設)",
      rowIndex: i + 1
    });
  }

  const allNames = Object.keys(byName);
  const totalMembers = allNames.reduce((sum, n) => sum + byName[n].length, 0);
  Logger.log(`📖 共 ${totalMembers} 位會友，${allNames.length} 個不同名字`);
  Logger.log('');

  // ── 1. 完全同名 ──
  const exactDuplicates = allNames.filter(n => byName[n].length > 1);
  Logger.log('═══════════════════════════════════════');
  Logger.log(`🔴 【完全相同的名字】共 ${exactDuplicates.length} 組`);
  Logger.log('═══════════════════════════════════════');
  if (exactDuplicates.length === 0) {
    Logger.log('  （無）');
  } else {
    exactDuplicates.forEach(name => {
      Logger.log(`「${name}」共 ${byName[name].length} 位：`);
      byName[name].forEach(m => {
        Logger.log(`    ▸ ${name}(${m.uid})  [${m.gender}, ${m.group}, ${m.role}]`);
      });
      Logger.log('');
    });
  }

  // ── 2. 高度相似（Levenshtein 距離 = 1）──
  Logger.log('═══════════════════════════════════════');
  Logger.log('🟡 【高度相似的名字】（字差 1，可能是同人或近親）');
  Logger.log('═══════════════════════════════════════');
  const similarPairs = [];
  for (let i = 0; i < allNames.length; i++) {
    for (let j = i + 1; j < allNames.length; j++) {
      const a = allNames[i], b = allNames[j];
      // 只比對長度差 ≤ 1 的（節省時間）
      if (Math.abs(a.length - b.length) > 1) continue;
      // 太短不比對（單字名容易誤判）
      if (a.length < 2 || b.length < 2) continue;
      const dist = _levenshtein(a, b);
      if (dist === 1) {
        similarPairs.push({ a: a, b: b });
      }
    }
  }
  if (similarPairs.length === 0) {
    Logger.log('  （無）');
  } else {
    similarPairs.forEach(p => {
      const aList = byName[p.a].map(m => `${p.a}(${m.uid})`).join(', ');
      const bList = byName[p.b].map(m => `${p.b}(${m.uid})`).join(', ');
      Logger.log(`「${p.a}」 ↔ 「${p.b}」`);
      Logger.log(`    ▸ ${aList}  [${byName[p.a][0].gender}, ${byName[p.a][0].group}]`);
      Logger.log(`    ▸ ${bList}  [${byName[p.b][0].gender}, ${byName[p.b][0].group}]`);
      Logger.log('');
    });
  }

  // ── 3. 共同前綴 ≥ 2 字（資訊用，可能是親屬）──
  Logger.log('═══════════════════════════════════════');
  Logger.log('🟢 【共同前綴 ≥ 2 字的同姓會友】（資訊用，多為親屬）');
  Logger.log('═══════════════════════════════════════');
  // 以前 2 字分組
  const byPrefix = {};
  allNames.forEach(n => {
    if (n.length < 3) return;  // 少於 3 字不分組（前 2 字 = 全名情況）
    const p = n.substring(0, 2);
    if (!byPrefix[p]) byPrefix[p] = [];
    byPrefix[p].push(n);
  });
  const prefixGroups = Object.keys(byPrefix).filter(p => byPrefix[p].length > 1);
  if (prefixGroups.length === 0) {
    Logger.log('  （無）');
  } else {
    prefixGroups.sort((a, b) => byPrefix[b].length - byPrefix[a].length);
    prefixGroups.forEach(p => {
      Logger.log(`「${p}__」共 ${byPrefix[p].length} 位：`);
      byPrefix[p].forEach(n => {
        const list = byName[n].map(m => `${n}(${m.uid})`).join(', ');
        Logger.log(`    ▸ ${list}`);
      });
      Logger.log('');
    });
  }

  Logger.log('═══════════════════════════════════════');
  Logger.log('✅ 掃描完成');
  Logger.log('═══════════════════════════════════════');

  return {
    totalMembers,
    uniqueNames: allNames.length,
    exactDuplicates: exactDuplicates.length,
    similarPairs: similarPairs.length,
    prefixGroups: prefixGroups.length
  };
}

// ═══════════════════════════════════════════════════════════
//  findHomophoneMembers
//
//  同音偵測：用拼音比對找出「不同字但唸法相同」的姓名
//  例：王小明 (wáng xiǎo míng) ↔ 王小銘 (wáng xiǎo míng)
//
//  拼音資料：UrlFetch 從 mozillazg/pinyin-data GitHub 抓
//           （只在 cache miss 時抓一次，之後永久用 cache）
//
//  執行：GAS 編輯器選 findHomophoneMembers → ▶ 執行
// ═══════════════════════════════════════════════════════════
function findHomophoneMembers() {
  Logger.log('🔍 開始同音字偵測（首次執行需 5-10 秒下載拼音資料）...');

  let pinyinMap;
  try {
    pinyinMap = _loadPinyinMap();
  } catch (e) {
    Logger.log('❌ 載入拼音資料失敗：' + e.message);
    return;
  }
  Logger.log(`📚 拼音資料已就緒（${Object.keys(pinyinMap).length} 字）`);

  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) { Logger.log('⚠️ 會友名單為空'); return; }

  // 以拼音為 key 分組：{ "wang-xiao-ming": [{ name, uid, ... }] }
  const byPinyin = {};
  let unmappedChars = new Set();

  for (let i = 1; i < data.length; i++) {
    const name = data[i][0] ? String(data[i][0]).trim() : "";
    if (!name) continue;

    const pyResult = _nameToPinyin(name, pinyinMap);
    if (pyResult.unmapped.length > 0) pyResult.unmapped.forEach(c => unmappedChars.add(c));
    const py = pyResult.pinyin;

    if (!byPinyin[py]) byPinyin[py] = [];
    byPinyin[py].push({
      name: name,
      uid:    data[i][7]  ? String(data[i][7]).trim()  : "(無編號)",
      gender: data[i][1]  ? String(data[i][1]).trim()  : "?",
      group:  data[i][9]  ? String(data[i][9]).trim()  : "(無組)",
      role:   data[i][10] ? String(data[i][10]).trim() : "(未設)"
    });
  }

  // 找「同拼音但不同字」的群組
  const homophones = [];
  Object.keys(byPinyin).forEach(py => {
    const list = byPinyin[py];
    const distinctNames = new Set(list.map(m => m.name));
    if (distinctNames.size > 1) homophones.push({ py, list, distinctNames });
  });

  Logger.log('═══════════════════════════════════════');
  Logger.log(`🟣 【同音字（不同寫法但唸法相同）】共 ${homophones.length} 組`);
  Logger.log('═══════════════════════════════════════');
  if (homophones.length === 0) {
    Logger.log('  （無）');
  } else {
    // 依組內人數降序排列
    homophones.sort((a, b) => b.list.length - a.list.length);
    homophones.forEach(h => {
      const namesStr = [...h.distinctNames].join(' / ');
      Logger.log(`「${namesStr}」  → 拼音：${h.py}`);
      h.list.forEach(m => {
        Logger.log(`    ▸ ${m.name}(${m.uid})  [${m.gender}, ${m.group}, ${m.role}]`);
      });
      Logger.log('');
    });
  }

  if (unmappedChars.size > 0) {
    Logger.log('───────────────────────────────────────');
    Logger.log(`⚠️ 拼音庫未涵蓋 ${unmappedChars.size} 個罕見字（這些字仍以原字比對，不影響其他結果）：`);
    Logger.log('  ' + [...unmappedChars].join(' '));
  }

  Logger.log('═══════════════════════════════════════');
  return { homophoneGroups: homophones.length, unmappedChars: [...unmappedChars] };
}

/**
 * 把姓名轉成拼音字串（用 - 連接）
 * 找不到的字保留原字（會與相同原字者匹配，但不會與真正同音者誤判）
 */
function _nameToPinyin(name, map) {
  const parts = [];
  const unmapped = [];
  for (let i = 0; i < name.length; i++) {
    const ch = name.charAt(i);
    if (map[ch]) {
      parts.push(map[ch]);
    } else {
      parts.push(ch);
      // 排除非中文字（英數、空白）
      if (/[一-鿿]/.test(ch)) unmapped.push(ch);
    }
  }
  return { pinyin: parts.join('-'), unmapped };
}

/**
 * 載入拼音資料（CacheService 永久保留；cache miss 才從 GitHub 抓）
 * 資料來源：mozillazg/pinyin-data 的 kMandarin.txt（CC0，可自由使用）
 */
function _loadPinyinMap() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'PINYIN_MAP_V1';

  // 先試讀（含分片支援）
  const single = cache.get(CACHE_KEY);
  if (single !== null) {
    try { return JSON.parse(single); } catch (e) { /* 損毀 */ }
  }
  const cntStr = cache.get(CACHE_KEY + '_CNT');
  if (cntStr !== null) {
    const count = parseInt(cntStr, 10);
    const keys = Array.from({ length: count }, (_, i) => CACHE_KEY + '_CHK_' + i);
    const parts = cache.getAll(keys);
    const assembled = keys.map(k => parts[k] || '').join('');
    if (assembled) {
      try { return JSON.parse(assembled); } catch (e) { /* 損毀 */ }
    }
  }

  // Cache miss → 從 GitHub 抓
  Logger.log('  ↻ 從 GitHub 下載拼音資料庫...');
  const url = 'https://raw.githubusercontent.com/mozillazg/pinyin-data/master/kMandarin.txt';
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('GitHub 拼音資料下載失敗 HTTP ' + resp.getResponseCode());
  }
  const text = resp.getContentText();

  const map = {};
  text.split('\n').forEach(line => {
    // 格式：U+4E00: yi1  # 一
    const m = line.match(/^U\+([0-9A-F]+):\s*([a-z]+)\d+/i);
    if (!m) return;
    const char = String.fromCodePoint(parseInt(m[1], 16));
    map[char] = m[2].toLowerCase();
  });

  // 寫入 cache（支援分片）
  const json = JSON.stringify(map);
  const CHUNK = 90000;
  if (json.length <= CHUNK) {
    cache.put(CACHE_KEY, json, 21600);
    cache.remove(CACHE_KEY + '_CNT');
  } else {
    const chunks = [];
    for (let i = 0; i < json.length; i += CHUNK) chunks.push(json.slice(i, i + CHUNK));
    const entries = {};
    entries[CACHE_KEY + '_CNT'] = String(chunks.length);
    chunks.forEach((c, i) => { entries[CACHE_KEY + '_CHK_' + i] = c; });
    cache.putAll(entries, 21600);
    cache.remove(CACHE_KEY);
  }
  Logger.log(`  ✅ 拼音資料已快取（${Object.keys(map).length} 字）`);
  return map;
}

/** Levenshtein distance（編輯距離）— 用於姓名相似度比對 */
function _levenshtein(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  const m = [];
  for (let i = 0; i <= b.length; i++) m[i] = [i];
  for (let j = 0; j <= a.length; j++) m[0][j] = j;
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        m[i][j] = m[i - 1][j - 1];
      } else {
        m[i][j] = Math.min(
          m[i - 1][j - 1] + 1,
          m[i][j - 1] + 1,
          m[i - 1][j] + 1
        );
      }
    }
  }
  return m[b.length][a.length];
}
