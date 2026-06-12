let adminGroupsList = [];
let currentEditingGroup = null; // 改為儲存完整的 group 物件（包含 uuid）
let verifiedAdminCode = ""; // 儲存已驗證的管理員代碼
let activeTab = 'regular'; // 'regular' or 'happy'
let globalIsAdmin = false; // 儲存最高管理員權限狀態

// showLoading / hideLoading / ensureAPIReady 由 config.js 提供。

window.onload = async () => {
    try {
        await ensureAPIReady();
    } catch (e) {
        userNotification.error("系統路由啟動失敗，請重新整理");
    }
    
    // 註冊 Enter 鍵登入管理
    const adminInput = document.getElementById('adminInput');
    if (adminInput) {
        adminInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                verifyAdmin();
            }
        });
    }
};

async function callAPI(action, data = {}) {
    if (typeof window.churchAPI !== 'function') throw new Error("安全路由尚未載入");
    return await window.churchAPI(action, data);
}

// 1. 驗證管理員身分
async function verifyAdmin() {
    const code = document.getElementById('adminInput').value.trim();
    if (!code) return userNotification.warning("請輸入代碼");

    showLoading("驗證權限中...");
    try {
        const res = await callAPI('findGroupByCode', { groupCode: code });

        if (res.success) {
            verifiedAdminCode = code;
            document.getElementById('login-panel').style.display = 'none';
            document.getElementById('manage-panel').style.display = 'block';
            await loadGroups();
        } else {
            userNotification.error("❌ 權限不足或代碼錯誤！" + (res.message ? " " + res.message : ""));
        }
    } catch (e) {
        userNotification.error("連線發生錯誤: " + e.message);
    } finally {
        hideLoading();
    }
}

// 2. 載入小組清單
async function loadGroups() {
    showLoading("正在撈取小組資料庫...");
    try {
        // ✅ 傳入已驗證的管理員代碼
        const res = await callAPI('getAdminGroupsList', { authCode: verifiedAdminCode });
        if (res.success) {
            adminGroupsList = res.groups;
            globalIsAdmin = res.isAdmin; // ✅ 儲存權限狀態
            updatePermissionBadge(res.isAdmin); // ✅ 更新權限徽章
            renderTable(res.isAdmin); // ✅ 傳入權限等級
        } else {
            userNotification.error("載入失敗：" + res.message);
        }
    } catch (e) {
        userNotification.error("連線發生錯誤: " + e.message);
    } finally {
        hideLoading();
    }
}

// 2.7 切換分頁
function switchTab(tabName) {
    activeTab = tabName;
    const regTab = document.getElementById('btn-tab-regular');
    const happyTab = document.getElementById('btn-tab-happy');
    if (regTab) regTab.classList.toggle('active', tabName === 'regular');
    if (happyTab) happyTab.classList.toggle('active', tabName === 'happy');
    renderTable(globalIsAdmin);
}

// 2.5 更新權限徽章顯示
function updatePermissionBadge(isAdmin) {
    const badge = document.getElementById('permission-badge');
    if (isAdmin) {
        badge.innerHTML = '👑 最高權限 - 系統管理員';
        badge.style.background = '#c62828';
    } else {
        badge.innerHTML = '👤 小組權限 - 小組管理員';
        badge.style.background = '#FF9800';
    }
}

// 3. 渲染表格
function renderTable(isAdmin) {
    const tbody = document.querySelector('#groupsTable tbody');
    const dateColumn = document.getElementById('dateColumn');
    
    // ✅ 根據權限動態調整欄位顯示
    if (isAdmin) {
        dateColumn.style.display = ''; // 顯示日期欄
    } else {
        dateColumn.style.display = 'none'; // 隱藏日期欄
    }
    
    const filteredGroups = adminGroupsList.filter(g => {
        if (activeTab === 'happy') {
            return g.type === '幸福小組';
        } else {
            return g.type !== '幸福小組';
        }
    });
    
    if (filteredGroups.length === 0) {
        const colspan = isAdmin ? '5' : '4';
        tbody.innerHTML = `<tr><td colspan="${colspan}" style="color: #999;">目前沒有小組資料</td></tr>`;
        return;
    }

    tbody.innerHTML = filteredGroups.map((g) => {
        // 找出在原始名單中的 index，以便傳給編輯 Modal
        const originalIndex = adminGroupsList.findIndex(orig => orig.uuid === g.uuid);

        // 狀態顯示樣式 (支援幸福小組結案狀態)
        let statusBadge = '';
        if (g.status === '顯示') {
            statusBadge = '<span style="background: #4CAF50; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px;">✅ 顯示</span>';
        } else if (g.status === '結案') {
            statusBadge = '<span style="background: #E91E63; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px;">🍀 結案</span>';
        } else {
            statusBadge = '<span style="background: #9E9E9E; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px;">🚫 隱藏</span>';
        }
        
        // 日期格式化
        const dateDisplay = g.date ? formatDate(g.date) : '-';
        
        // ✅ 根據權限決定顯示的欄位
        const dateCell = isAdmin ? `<td style="color: #666; font-size: 14px;">${dateDisplay}</td>` : '';
        
        // 幸福小組標示
        const groupTypeBadge = g.type === '幸福小組'
            ? ' <span class="role-badge role-best" style="font-size:11px; padding:2px 6px; width:auto; border-radius:4px; vertical-align:middle; background: #fff3e0; color: #ef6c00;">幸福</span>'
            : '';

        // 結案按鈕 (限管理員且未結案)
        const concludeBtn = (isAdmin && g.status !== '結案')
            ? `<button class="btn" style="background: #e91e63; padding: 6px 12px; font-size: 13px; margin-left: 8px;" onclick="concludeGroup('${g.name}')">🍀 結案</button>`
            : '';

        // 徹底刪除按鈕 (限管理員且已結案)
        const deleteBtn = (isAdmin && g.status === '結案')
            ? `<button class="btn" style="background: #f44336; padding: 6px 12px; font-size: 13px; margin-left: 8px;" onclick="deleteGroup('${g.name}')">🗑️ 徹底刪除</button>`
            : '';
        
        return `
            <tr>
                <td style="font-weight: bold; font-size: 16px;">${g.name}${groupTypeBadge}</td>
                <td><span style="background: #eee; padding: 4px 10px; border-radius: 12px; font-family: monospace;">${g.code}</span></td>
                <td>${statusBadge}</td>
                ${dateCell}
                <td>
                    <button class="btn" style="background: #2196F3; padding: 6px 12px; font-size: 13px;" onclick="openEditModal(${originalIndex})">✏️ 編輯</button>${concludeBtn}${deleteBtn}
                </td>
            </tr>
        `;
    }).join('');
}

// 3.5 日期格式化：優先用 config.js 的 formatYMD，無法解析時顯示原值
function formatDate(dateValue) {
    if (!dateValue) return '-';
    return window.formatYMD(dateValue) || String(dateValue);
}

// 4. 開啟編輯彈窗
function openEditModal(index) {
    const group = adminGroupsList[index];
    currentEditingGroup = group; // ✅ 儲存完整的 group 物件（包含 uuid）
    
    document.getElementById('editOldName').value = group.name;
    document.getElementById('editNewName').value = group.name;
    document.getElementById('editNewCode').value = group.code;
    
    const statusSelect = document.getElementById('editNewStatus');
    // 先移除之前動態補上的結案選項
    const concludeOpt = statusSelect.querySelector('option[value="結案"]');
    if (concludeOpt) concludeOpt.remove();

    // 如果該組已結案，動態加入唯讀結案選項並停用下拉選單
    if (group.status === '結案') {
        const opt = document.createElement('option');
        opt.value = '結案';
        opt.innerHTML = '🍀 結案 (已封存，不可修改狀態)';
        statusSelect.appendChild(opt);
        statusSelect.value = '結案';
        statusSelect.disabled = true;
    } else {
        statusSelect.value = group.status || '顯示';
        statusSelect.disabled = false;
    }
    
    document.getElementById('edit-group-modal').style.display = 'block';
}

function closeEditModal() {
    document.getElementById('edit-group-modal').style.display = 'none';
}

// 5. 儲存修改
async function saveGroupEdit() {
    const newName = document.getElementById('editNewName').value.trim();
    const newCode = document.getElementById('editNewCode').value.trim();
    const newStatus = document.getElementById('editNewStatus').value;

    if (!newName) return userNotification.warning("名稱不可為空！");
    if (newCode.length < 4) return userNotification.warning("代碼至少需要 4 碼！");

    const hasChanges =
        newName !== currentEditingGroup.name ||
        newCode !== currentEditingGroup.code ||
        newStatus !== currentEditingGroup.status;

    if (!hasChanges) return userNotification.warning("⚠️ 您沒有做任何修改！");

    if (newName !== currentEditingGroup.name) {
        if (!confirm(`⚠️ 警告：您即將把【${currentEditingGroup.name}】改名為【${newName}】\n\n系統將會同步重新命名資料庫中的分頁，此動作需要幾秒鐘，確定要執行嗎？`)) {
            return;
        }
    }

    showLoading("正在更新資料庫與同步分頁名稱，請勿關閉網頁...");
    try {
        const res = await callAPI('updateGroupInfo', { 
            uuid: currentEditingGroup.uuid,
            oldName: currentEditingGroup.name,
            newName: newName, 
            newCode: newCode,
            newStatus: newStatus
        });

        if (res.success) {
            userNotification.success('✅ 修改成功！分頁名稱已同步更新。');
            closeEditModal();
            if (newCode !== currentEditingGroup.code) {
                verifiedAdminCode = newCode;
            }
            await loadGroups();
        } else {
            userNotification.error('❌ 修改失敗：' + res.message);
        }
    } catch (e) {
        userNotification.error("連線發生錯誤: " + e.message);
    } finally {
        hideLoading();
    }
}

async function concludeGroup(groupName) {
    if (!confirm(`🍀 確定要將小組【${groupName}】結案嗎？\n\n此動作將會進行雲端封存，且會移除所有同工與成員的此小組關聯（但保留其在會友名單的名字），確定要執行嗎？`)) {
        return;
    }
    showLoading(`正在將小組【${groupName}】結案中...`);
    try {
        const res = await callAPI('happyGroup_conclude', { 
            groupName: groupName, 
            bestToUpgrade: [], 
            authCode: verifiedAdminCode 
        });
        if (res.success) {
            userNotification.success('✅ 小組已成功結案！');
            await loadGroups();
        } else {
            userNotification.error('❌ 結案失敗：' + res.message);
        }
    } catch (e) {
        userNotification.error("連線發生錯誤: " + e.message);
    } finally {
        hideLoading();
    }
}

async function deleteGroup(groupName) {
    if (!confirm(`⚠️ 警告：確定要徹底刪除小組【${groupName}】嗎？\n\n此動作將永久刪除該組在試算表中的所有分頁（名單與點名紀錄），且不可復原！`)) {
        return;
    }
    showLoading(`正在徹底刪除【${groupName}】...`);
    try {
        const res = await callAPI('happyGroup_delete', { groupName: groupName, authCode: verifiedAdminCode });
        if (res.success) {
            userNotification.success('✅ 小組已被徹底刪除！');
            await loadGroups();
        } else {
            userNotification.error('❌ 刪除失敗：' + res.message);
        }
    } catch (e) {
        userNotification.error("連線發生錯誤: " + e.message);
    } finally {
        hideLoading();
    }
}
