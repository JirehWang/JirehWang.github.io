/**
 * ============================================================
 *  🏰 牧區與小組群分類管理 — GAS 後端擴充 (簡化單表版)
 *  檔案名稱：Hierarchy.js
 * ============================================================
 */

// ============================================================
//  🛠️ 自動檢查與初始化 — 在 小組清單 補上 district 與 cluster 欄位
// ============================================================
function initHierarchySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var groupsSheet = ss.getSheetByName('小組清單');
  if (!groupsSheet) {
    Logger.log('❌ 找不到 小組清單 工作表');
    return;
  }

  var headers = groupsSheet.getRange(1, 1, 1, groupsSheet.getLastColumn()).getValues()[0];
  var hasDistrict = headers.indexOf('district') !== -1;
  var hasCluster  = headers.indexOf('cluster') !== -1;

  if (!hasDistrict) {
    var nextCol = groupsSheet.getLastColumn() + 1;
    groupsSheet.getRange(1, nextCol).setValue('district').setFontWeight('bold');
    Logger.log('✅ 小組清單 新增 district 欄位');
  }
  if (!hasCluster) {
    var nextCol = groupsSheet.getLastColumn() + 1;
    groupsSheet.getRange(1, nextCol).setValue('cluster').setFontWeight('bold');
    Logger.log('✅ 小組清單 新增 cluster 欄位');
  }
}

// ============================================================
//  📦 共用工具函數
// ============================================================

function _getSheetRows(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return data.map(function(row) {
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    return obj;
  });
}

function _getSheetHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function _getColIndex(sheet, colName) {
  var headers = _getSheetHeaders(sheet);
  var idx = headers.indexOf(colName);
  return idx === -1 ? -1 : idx + 1;
}

function _findRowByUuid(rows, uuid) {
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].uuid === uuid) return i;
  }
  return -1;
}

// ============================================================
//  🎯 主路由入口：handleHierarchyAction
// ============================================================
function handleHierarchyAction(action, data) {
  // 自動檢查並初始化
  try {
    initHierarchySheets();
  } catch (e) {
    Logger.log('⚠️ 自動初始化 Hierarchy 欄位失敗: ' + e.message);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var groupsSheet = ss.getSheetByName('小組清單');

  switch (action) {
    case 'getDistrictsAndClusters':
      return _handleGetDistrictsAndClusters(data, groupsSheet);
    case 'createDistrict':
      return _handleCreateDistrict(data, groupsSheet);
    case 'createGroupCluster':
      return _handleCreateGroupCluster(data, groupsSheet);
    case 'updateClusterGroups':
      return _handleUpdateClusterGroups(data, groupsSheet);
    default:
      return { success: false, message: '未知的 hierarchy action: ' + action };
  }
}

// ============================================================
//  1️⃣ getDistrictsAndClusters — 動態去重取得最新牧區與小組群列表
// ============================================================
function _handleGetDistrictsAndClusters(data, groupsSheet) {
  var groups = _getSheetRows(groupsSheet);
  
  var districtSet = {};
  var clusterSet = {};
  
  groups.forEach(function(g) {
    var dist = String(g.district || '').trim();
    var clust = String(g.cluster || '').trim();
    
    if (dist) {
      districtSet[dist] = true;
    }
    if (clust) {
      clusterSet[clust] = dist; // 記錄小組群所屬的牧區
    }
  });
  
  var districts = Object.keys(districtSet).map(function(name) {
    return { uuid: name, name: name };
  });
  
  var clusters = Object.keys(clusterSet).map(function(name) {
    return {
      uuid: name,
      name: name,
      districtUuid: clusterSet[name],
      districtName: clusterSet[name]
    };
  });
  
  var formattedGroups = groups.map(function(g) {
    return {
      uuid: g.uuid,
      name: g.name,
      code: g.code,
      type: g.type || '',
      status: g.status || '顯示',
      districtUuid: g.district || '',
      districtName: g.district || '',
      clusterUuid: g.cluster || '',
      clusterName: g.cluster || ''
    };
  });

  var ADMIN_CODE = _getAdminCode();
  var authCode = data.authCode;
  var isAdmin = (authCode === ADMIN_CODE);
  var matchedGroup = null;

  if (!isAdmin && authCode) {
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].code === authCode) {
        matchedGroup = groups[i];
        break;
      }
    }
  }

  return {
    success: true,
    isAdmin: isAdmin,
    groupName: matchedGroup ? matchedGroup.name : null,
    clusterUuid: matchedGroup ? (matchedGroup.cluster || null) : null,
    clusterName: matchedGroup ? (matchedGroup.cluster || null) : null,
    districts: districts,
    clusters: clusters,
    groups: formattedGroups
  };
}

// ============================================================
//  2️⃣ createDistrict — 建立新牧區 (管理員專屬)
// ============================================================
function _handleCreateDistrict(data, groupsSheet) {
  var ADMIN_CODE = _getAdminCode();
  if (data.authCode !== ADMIN_CODE) return { success: false, message: '無此操作權限！' };

  var name = (data.name || '').trim();
  if (!name) return { success: false, message: '牧區名稱不可為空' };

  var clusterNames = data.clusterUuids || []; // 前端選取的小組群名稱列表

  if (clusterNames.length > 0) {
    var groupRows = _getSheetRows(groupsSheet);
    var distCol = _getColIndex(groupsSheet, 'district');
    for (var i = 0; i < groupRows.length; i++) {
      if (clusterNames.indexOf(groupRows[i].cluster) !== -1) {
        groupsSheet.getRange(i + 2, distCol).setValue(name);
      }
    }
  }
  return { success: true, message: '牧區建立成功！' };
}

// ============================================================
//  3️⃣ createGroupCluster — 建立新小組群 (管理員/小組長)
// ============================================================
function _handleCreateGroupCluster(data, groupsSheet) {
  var authCode = data.authCode;
  if (!authCode) return { success: false, message: '缺少驗證代碼' };

  // 驗證權限 (管理員或小組長皆可)
  var ADMIN_CODE = _getAdminCode();
  var isAdmin = (authCode === ADMIN_CODE);
  if (!isAdmin) {
    var groups = _getSheetRows(groupsSheet);
    var found = false;
    for (var k = 0; k < groups.length; k++) {
      if (groups[k].code === authCode) { found = true; break; }
    }
    if (!found) return { success: false, message: '驗證代碼錯誤！' };
  }

  var name = (data.name || '').trim();
  if (!name) return { success: false, message: '小組群名稱不可為空' };

  var districtName = data.districtUuid || ''; // 歸屬的牧區名稱
  var groupUuids = data.groupUuids || [];

  if (groupUuids.length > 0) {
    var groupRows = _getSheetRows(groupsSheet);
    var clusterCol = _getColIndex(groupsSheet, 'cluster');
    var distCol = _getColIndex(groupsSheet, 'district');

    for (var i = 0; i < groupRows.length; i++) {
      if (groupUuids.indexOf(groupRows[i].uuid) !== -1) {
        groupsSheet.getRange(i + 2, clusterCol).setValue(name);
        if (districtName) {
          groupsSheet.getRange(i + 2, distCol).setValue(districtName);
        }
      }
    }
  }
  return { success: true, message: '小組群建立成功！' };
}

// ============================================================
//  4️⃣ updateClusterGroups — 更新小組群旗下小組 (移出/加入)
// ============================================================
function _handleUpdateClusterGroups(data, groupsSheet) {
  var authCode = data.authCode;
  if (!authCode) return { success: false, message: '缺少驗證代碼' };

  var clusterName = data.clusterUuid; // 目標小組群名稱
  var targetGroupUuids = data.groupUuids || [];
  if (!clusterName) return { success: false, message: '缺少小組群名稱' };

  var groupRows = _getSheetRows(groupsSheet);
  var clusterCol = _getColIndex(groupsSheet, 'cluster');
  var distCol = _getColIndex(groupsSheet, 'district');

  // 找到目標小組群原本擁有的牧區
  var parentDistrictName = '';
  for (var i = 0; i < groupRows.length; i++) {
    if (groupRows[i].cluster === clusterName && groupRows[i].district) {
      parentDistrictName = groupRows[i].district;
      break;
    }
  }

  for (var j = 0; j < groupRows.length; j++) {
    var g = groupRows[j];
    var isCurrentlyIn = (g.cluster === clusterName);
    var shouldBeIn = (targetGroupUuids.indexOf(g.uuid) !== -1);

    if (isCurrentlyIn && !shouldBeIn) {
      // ❌ 移出群組：清空 cluster 與 district
      groupsSheet.getRange(j + 2, clusterCol).setValue('');
      groupsSheet.getRange(j + 2, distCol).setValue('');
    } else if (!isCurrentlyIn && shouldBeIn) {
      // ➕ 拉入群組：設定 cluster，district 自動繼承
      groupsSheet.getRange(j + 2, clusterCol).setValue(clusterName);
      groupsSheet.getRange(j + 2, distCol).setValue(parentDistrictName);
    }
  }

  return { success: true, message: '小組群成員更新成功！' };
}

// ============================================================
//  🔧 enrichAdminGroupsListWithHierarchy
// ============================================================
function enrichAdminGroupsListWithHierarchy(result, authCode) {
  if (!result || !result.success) return result;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var groupsSheet = ss.getSheetByName('小組清單');
  var groups = groupsSheet ? _getSheetRows(groupsSheet) : [];

  var districtSet = {};
  var clusterSet = {};

  groups.forEach(function(g) {
    var dist = String(g.district || '').trim();
    var clust = String(g.cluster || '').trim();
    if (dist) districtSet[dist] = true;
    if (clust) clusterSet[clust] = dist;
  });

  var districts = Object.keys(districtSet).map(function(name) {
    return { uuid: name, name: name };
  });

  var clusters = Object.keys(clusterSet).map(function(name) {
    return {
      uuid: name,
      name: name,
      districtUuid: clusterSet[name],
      districtName: clusterSet[name]
    };
  });

  // 為每個 group 補上歸屬名稱與屬性
  var groupMap = {};
  groups.forEach(function(g) {
    groupMap[g.uuid] = { district: g.district || '', cluster: g.cluster || '' };
  });

  if (result.groups && Array.isArray(result.groups)) {
    result.groups = result.groups.map(function(g) {
      var hierarchy = groupMap[g.uuid] || {};
      g.districtUuid = hierarchy.district || '';
      g.districtName = hierarchy.district || '';
      g.clusterUuid = hierarchy.cluster || '';
      g.clusterName = hierarchy.cluster || '';
      return g;
    });
  }

  result.districts = districts;
  result.clusters = clusters;

  // 如果呼叫者不是管理員，補上該小組長所屬的小組群資訊
  var ADMIN_CODE = _getAdminCode();
  if (authCode && authCode !== ADMIN_CODE) {
    var matchedGroup = groups.find(function(g) { return g.code === authCode; });
    if (matchedGroup) {
      result.groupName = matchedGroup.name;
      result.clusterUuid = matchedGroup.cluster || '';
      result.clusterName = matchedGroup.cluster || '';
    }
  }

  return result;
}

// ============================================================
//  🔧 writeGroupHierarchyFields
// ============================================================
function writeGroupHierarchyFields(groupUuid, districtName, clusterName) {
  if (districtName === undefined && clusterName === undefined) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var groupsSheet = ss.getSheetByName('小組清單');
  var groupRows = _getSheetRows(groupsSheet);
  var idx = _findRowByUuid(groupRows, groupUuid);
  if (idx === -1) return;

  if (districtName !== undefined) {
    var distCol = _getColIndex(groupsSheet, 'district');
    if (distCol > 0) groupsSheet.getRange(idx + 2, distCol).setValue(districtName || '');
  }
  if (clusterName !== undefined) {
    var clusterCol = _getColIndex(groupsSheet, 'cluster');
    if (clusterCol > 0) groupsSheet.getRange(idx + 2, clusterCol).setValue(clusterName || '');
  }
}

// ============================================================
//  🔧 assignNewGroupToCluster
// ============================================================
function assignNewGroupToCluster(newGroupUuid, targetClusterName, newClusterName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var groupsSheet = ss.getSheetByName('小組清單');

  var clusterCol = _getColIndex(groupsSheet, 'cluster');
  var distCol = _getColIndex(groupsSheet, 'district');
  var groupRows = _getSheetRows(groupsSheet);
  var groupIdx = _findRowByUuid(groupRows, newGroupUuid);
  if (groupIdx === -1) return;

  var finalCluster = targetClusterName || newClusterName || '';
  if (finalCluster) {
    groupsSheet.getRange(groupIdx + 2, clusterCol).setValue(finalCluster);

    if (targetClusterName) {
      var parentDistrictName = '';
      for (var i = 0; i < groupRows.length; i++) {
        if (groupRows[i].cluster === targetClusterName && groupRows[i].district) {
          parentDistrictName = groupRows[i].district;
          break;
        }
      }
      if (parentDistrictName) {
        groupsSheet.getRange(groupIdx + 2, distCol).setValue(parentDistrictName);
      }
    }
  }
}

// ============================================================
//  🔐 管理員代碼讀取
// ============================================================
function _getAdminCode() {
  if (typeof ADMIN_CODE !== 'undefined') {
    return ADMIN_CODE;
  }
  return 'LK31';
}
