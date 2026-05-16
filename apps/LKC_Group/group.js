// --- 基礎變數設定 ---
const urlParams = new URLSearchParams(window.location.search);
const groupName = urlParams.get('name');
const groupCode = urlParams.get('code');

let currentMembers = []; // [{ name, uid, role }]
let editingMembers = [];
let recentRecordsData = [];
let nameDirectory = {};  // uid → name 反查表（從後端 RAW_MODE 回傳）

// showLoading / hideLoading / ensureAPIReady 由 config.js 提供。

// 共用：UID 反查姓名（找不到就直接回傳原字串）
function resolveDisplayName(uidOrName) {
  if (!uidOrName) return "";
  const s = String(uidOrName).trim();
  if (/^LK\d+$/i.test(s)) return nameDirectory[s.toUpperCase()] || s;
  return s; // 不是 UID 格式（可能是新朋友）就直接顯示
}

// --- 📦 網頁載入啟動流程 ---
window.onload = async () => {
    try {
        showLoading("🚀 正在啟動安全通道...");
        await ensureAPIReady();
        document.getElementById('displayGroupName').innerText = groupName || '未知名組別';
        document.getElementById('attendanceDate').valueAsDate = new Date();
        await checkGroupStatus();
    } catch (e) {
        console.error(e);
        userNotification.error("系統啟動失敗：" + e.message);
        hideLoading();
    }
};

async function callAPI(action, data = {}) {
    if (typeof window.churchAPI !== 'function') throw new Error("安全路由尚未載入");
    return await window.churchAPI(action, data);
}

// 💡 更新：加入陪伴同工的 CSS Class
function getRoleClass(role) {
    if (role === '核心同工') return 'role-core';
    if (role === '一般同工') return 'role-general';
    if (role === '小羊') return 'role-sheep';
    if (role === '陪伴同工') return 'role-companion'; 
    return 'role-default';
}

function normalizeNames(inputString) {
    if (!inputString) return "";
    const splitRegex = /[^\u4e00-\u9fa5a-zA-Z0-9\s]+/; 
    return inputString.split(splitRegex).map(s => s.trim()).filter(n => n).join(',');
}

async function checkGroupStatus() {
    showLoading("正在載入點名單與聚會紀錄...");
    try {
        const res = await callAPI('checkGroupStatus', { groupName });
        if (res.isInitialized) {
            currentMembers = res.members;
            document.getElementById('attendance-panel').style.display = 'block';
            document.getElementById('init-panel').style.display = 'none'; 
            renderMemberList(res.members);
            
            if (groupCode) {
                document.getElementById('scheduleBtn').style.display = 'inline-block';
                await loadGroupProgress();
            }
        } else {
            document.getElementById('init-panel').style.display = 'block';
            document.getElementById('attendance-panel').style.display = 'none'; 
        }
    } catch (e) {
        userNotification.error("載入失敗，請重新整理頁面。");
    } finally {
        hideLoading();
    }
}

function goToSchedule() {
    if (!groupCode) return userNotification.warning("未取得小組編號，無法跳轉。");
    window.open(`https://jirehwang.github.io/LKC1958_June_1.github.io/?id=${groupCode}`, '_blank');
}

function goToFullStats() {
    window.open(`https://jirehwang.github.io/LKGroup.github.io/stats.html?id=${groupCode || ''}`, '_blank');
}

async function loadGroupProgress() {
    const tbody = document.getElementById('progressTableBody');
    if (!tbody) return;
    document.getElementById('progressSection').style.display = 'block';

    try {
        const res = await callAPI('getStats', { 
            groupName: groupName, 
            groupCode: groupCode, 
            startDate: "RAW_MODE", 
            endDate: "" 
        });
        
        if (res.success && res.data.length > 0) {
            // \u63a5\u6536 nameDirectory \u7528\u65bc UID \u53cd\u67e5
            if (res.nameDirectory) nameDirectory = res.nameDirectory;

            const sortedData = res.data.slice().sort((a, b) => {
                return new Date(b[0]).getTime() - new Date(a[0]).getTime();
            });
            recentRecordsData = sortedData.slice(0, 3);

            const splitRegex = /[^\u4e00-\u9fa5a-zA-Z0-9\s]+/;

            tbody.innerHTML = recentRecordsData.map((row, index) => {
                const dateObj = new Date(row[0]);
                const dateStr = `${dateObj.getMonth() + 1}/${dateObj.getDate()}`;

                const fullDateStr = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-${String(dateObj.getDate()).padStart(2, '0')}`;
                row.fullDateStr = fullDateStr;

                const presentRaw = row[1] ? row[1].toString().split(splitRegex).map(s=>s.trim()).filter(n=>n) : [];
                const present = presentRaw.map(resolveDisplayName);  // UID \u2192 \u59d3\u540d
                const newFriends = row[3] ? row[3].toString().split(splitRegex).map(s=>s.trim()).filter(n=>n) : [];
                const totalCount = present.length + newFriends.length;
                
                let namesHtml = present.join('、');
                if (newFriends.length > 0) {
                    namesHtml += ` <span style="color:#ef6c00; font-size:12px; font-weight:bold;">(+新朋友: ${newFriends.join('、')})</span>`;
                }

                return `
                    <tr>
                        <td><span style="background: #e3f2fd; padding: 3px 8px; border-radius: 12px; font-weight: bold; font-size: 12px;">${dateStr}</span></td>
                        <td style="font-weight: bold; font-size: 16px;">${totalCount} 人</td>
                        <td style="text-align: left; font-size: 13px; color: #555;">
                            ${namesHtml || '無出席紀錄'}
                            <button class="edit-record-btn" onclick="openEditAttendanceModal(${index})">✏️</button>
                        </td>
                    </tr>
                `;
            }).join('');
        } else {
            tbody.innerHTML = '<tr><td colspan="3" style="color: #999; padding: 20px;">目前尚無聚會紀錄</td></tr>';
        }
    } catch (e) {
        tbody.innerHTML = '<tr><td colspan="3" style="color: red;">讀取紀錄失敗</td></tr>';
    }
}

// 💡 更新：歷史紀錄修改時，同樣照常顯示所有人
function openEditAttendanceModal(index) {
    const row = recentRecordsData[index];
    const originalDate = row.fullDateStr;
    const splitRegex = /[^\u4e00-\u9fa5a-zA-Z0-9\s]+/; 
    const presentUidArr = row[1] ? row[1].toString().split(splitRegex).map(s=>s.trim()).filter(n=>n) : [];
    const presentUidSet = new Set(presentUidArr.map(s => s.toUpperCase()));
    const newFriendsStr = row[3] ? row[3].toString().split(splitRegex).map(s=>s.trim()).filter(n=>n).join(',') : '';

    document.getElementById('editOriginalDate').value = originalDate;
    document.getElementById('editAttendanceDate').value = originalDate;
    document.getElementById('editNewFriends').value = newFriendsStr;

    const listDiv = document.getElementById('editAttendanceMemberList');

    // checkbox value 用 UID
    listDiv.innerHTML = currentMembers.map(m => {
        const uid = m.uid || '';
        const isChecked = uid && presentUidSet.has(uid.toUpperCase()) ? 'checked' : '';
        const roleClass = getRoleClass(m.role);
        return `
            <div class="member-item">
                <input type="checkbox" class="edit-attendance-check" value="${uid}" data-name="${m.name}" ${isChecked}>
                <span class="role-badge ${roleClass}">${m.role}</span>
                <span style="font-size: 16px; font-weight: bold; color: #333;">${m.name}</span>
            </div>
        `;
    }).join('');

    document.getElementById('edit-attendance-modal').style.display = 'block';
}

function closeEditAttendanceModal() {
    document.getElementById('edit-attendance-modal').style.display = 'none';
}

async function submitAttendanceEdit() {
    const originalDate = document.getElementById('editOriginalDate').value;
    const newDate = document.getElementById('editAttendanceDate').value;
    const newFriends = normalizeNames(document.getElementById('editNewFriends').value);

    // 都送 UID（cb.value 已是 UID）；篩掉空字串
    const present = Array.from(document.querySelectorAll('.edit-attendance-check:checked')).map(cb => cb.value).filter(v => v);
    const absent = Array.from(document.querySelectorAll('.edit-attendance-check:not(:checked)')).map(cb => cb.value).filter(v => v);

    if (present.length === 0 && !newFriends) {
        if (!confirm("修改後出席人數為 0，確定要儲存嗎？")) return;
    }

    showLoading("正在更新點名紀錄...");
    try {
        const res = await callAPI('updateAttendanceRecord', { groupName, originalDate, newDate, present, absent, newFriends });
        if (res.success) {
            userNotification.success('修改成功！');
            closeEditAttendanceModal();
            if (groupCode) await loadGroupProgress();
        } else { userNotification.error('修改失敗：' + res.message); }
    } finally { hideLoading(); }
}

async function deleteAttendanceRecord() {
    const originalDate = document.getElementById('editOriginalDate').value;
    if (!confirm(`確定要將【${originalDate}】的點名紀錄完全刪除嗎？刪除後無法復原喔！`)) return;

    showLoading("正在刪除紀錄...");
    try {
        const res = await callAPI('deleteAttendanceRecord', { groupName, originalDate });
        if (res.success) {
            userNotification.success('紀錄已刪除！');
            closeEditAttendanceModal();
            if (groupCode) await loadGroupProgress();
        } else { userNotification.error('刪除失敗：' + res.message); }
    } finally { hideLoading(); }
}

async function initGroup() {
    const rawMembers = document.getElementById('memberInput').value.split('\n').filter(n => n.trim());
    if (rawMembers.length === 0) return userNotification.warning('請輸入名單');
    const members = rawMembers.map(name => ({ name: name.trim(), role: '小羊' }));

    showLoading("正在建立雲端分頁，這可能需要幾秒鐘...");
    try {
        const res = await callAPI('initGroup', { groupName, members });
        if (res.success) {
            await checkGroupStatus();
        }
        else { userNotification.error(res.message); }
    } finally { hideLoading(); }
}

// 💡 更新：今日點名時，照常顯示所有組員 (包含陪伴同工)，讓他們能被記錄
//    checkbox value 用 UID（後端用 UID 儲存），data-name 留作備援顯示
function renderMemberList(members) {
    const list = document.getElementById('memberList');
    if (members.length === 0) {
        list.innerHTML = '<div style="color:#999; padding: 10px;">目前名單為空</div>';
        return;
    }
    list.innerHTML = members.map(m => {
        const roleClass = getRoleClass(m.role);
        const uid = m.uid || '';
        return `
            <div class="member-item">
                <input type="checkbox" class="attendance-check" value="${uid}" data-name="${m.name}">
                <span class="role-badge ${roleClass}">${m.role}</span>
                <span style="font-size: 16px; font-weight: bold; color: #333;">${m.name}</span>
            </div>
        `;
    }).join('');
}

function toggleEditMode() {
    const modal = document.getElementById('edit-modal');
    if (modal.style.display === 'block') {
        modal.style.display = 'none';
    } else {
        editingMembers = currentMembers.map(m => ({...m}));
        document.getElementById('newMemberInput').value = ""; 
        renderEditList();
        modal.style.display = 'block';
    }
}

// 💡 更新：編輯名單可拖曳排序，順序會套用到小組點名介面
//    為了避免拖曳後 index 失準，所有事件改用 data-name 對應
let _editSortableInstance = null;
function renderEditList() {
    const container = document.getElementById('editMemberList');
    if (editingMembers.length === 0) {
        container.innerHTML = '<div class="empty-hint">目前名單為空</div>';
        if (_editSortableInstance) { _editSortableInstance.destroy(); _editSortableInstance = null; }
        return;
    }
    container.innerHTML = editingMembers.map(m => {
        const safeName = (m.name || '').replace(/'/g, "&#39;");
        return `
            <div class="edit-member-item" data-name="${safeName}">
                <div style="display: flex; align-items: center; gap: 8px; flex: 1;">
                    <span class="drag-handle" title="按住拖曳排序"
                          style="cursor: grab; color:#999; font-size:18px; padding:0 6px; user-select:none; touch-action:none;">⋮⋮</span>
                    <span style="font-weight:bold; width: 60px; overflow: hidden; text-overflow: ellipsis;">${m.name}</span>
                    <select class="edit-role-select" onchange="updateMemberRoleByName('${safeName}', this.value)">
                        <option value="核心同工" ${m.role==='核心同工'?'selected':''}>核心同工</option>
                        <option value="一般同工" ${m.role==='一般同工'?'selected':''}>一般同工</option>
                        <option value="小羊" ${m.role==='小羊'?'selected':''}>小羊</option>
                        <option value="陪伴同工" ${m.role==='陪伴同工'?'selected':''}>陪伴同工</option>
                    </select>
                </div>
                <button class="btn-remove" onclick="removeEditMemberByName('${safeName}')">🗑️ 移除</button>
            </div>
        `;
    }).join('');

    // 初始化拖曳（每次 render 都重建）
    if (_editSortableInstance) { _editSortableInstance.destroy(); }
    if (typeof Sortable !== 'undefined') {
        _editSortableInstance = Sortable.create(container, {
            animation: 200,
            handle: '.drag-handle',
            ghostClass: 'sortable-ghost',
            chosenClass: 'sortable-chosen',
            onEnd: function() {
                // 依 DOM 當前順序重排 editingMembers
                const newOrder = Array.from(container.children)
                    .map(el => el.dataset.name)
                    .filter(n => n);
                editingMembers.sort((a, b) => {
                    const ia = newOrder.indexOf(a.name);
                    const ib = newOrder.indexOf(b.name);
                    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
                });
            }
        });
    }
}

// 💡 更新：新增名單的下拉選單（因為 HTML 寫死了，這裡用 JS 動態替換原本的結構）
// (請確保 HTML 裡面的 #newMemberRole 選單也有加上 <option value="陪伴同工">陪伴同工</option>)
// 下方是 JS 新增邏輯：
function addEditMember() {
    const input = document.getElementById('newMemberInput');
    let roleSelect = document.getElementById('newMemberRole');
    
    // 防呆：如果前端 HTML 還沒加，強制抓取當前值或預設給小羊
    const newName = input.value.trim();
    const newRole = roleSelect ? roleSelect.value : '小羊'; 
    
    if (!newName) return userNotification.warning("請輸入要新增的姓名！");
    if (editingMembers.some(m => m.name === newName)) return userNotification.warning("此人已經在名單中了！");

    editingMembers.push({ name: newName, role: newRole });
    input.value = ""; 
    renderEditList(); 
}

// 舊版（保留向下相容，依 index）
function updateMemberRole(index, newRole) { if (editingMembers[index]) editingMembers[index].role = newRole; }
function removeEditMember(index) {
    const nameToRemove = editingMembers[index] && editingMembers[index].name;
    if (!nameToRemove) return;
    if (confirm(`確定要將【${nameToRemove}】從名單中移除嗎？`)) {
        editingMembers.splice(index, 1);
        renderEditList();
    }
}

// 新版（不受拖曳影響，依姓名查找）
function updateMemberRoleByName(name, newRole) {
    const m = editingMembers.find(x => x.name === name);
    if (m) m.role = newRole;
}
function removeEditMemberByName(name) {
    if (confirm(`確定要將【${name}】從名單中移除嗎？`)) {
        editingMembers = editingMembers.filter(m => m.name !== name);
        renderEditList();
    }
}

async function saveUpdatedList() {
    if (editingMembers.length === 0) {
        if (!confirm('目前名單為空，確定要清空整個小組名單嗎？')) return;
    } else {
        if (!confirm('確定要儲存這份新名單嗎？')) return;
    }
    showLoading("正在更新雲端名單...");
    try {
        const res = await callAPI('updateMemberList', { groupName, members: editingMembers });
        if (res.success) {
            userNotification.success('名單更新成功！');
            currentMembers = [...editingMembers];
            renderMemberList(currentMembers);
            toggleEditMode();
        }
        else { userNotification.error('更新失敗：' + res.message); }
    } catch (e) { userNotification.error("連線發生錯誤，請稍後再試。"); } finally { hideLoading(); }
}

async function submitAttendance() {
    const date = document.getElementById('attendanceDate').value;
    // 送 UID 列表（cb.value 已是 UID）；篩掉空字串避免主日尚未綁定 UID 的會友造成髒資料
    const present = Array.from(document.querySelectorAll('.attendance-check:checked')).map(cb => cb.value).filter(v => v);
    const absent = Array.from(document.querySelectorAll('.attendance-check:not(:checked)')).map(cb => cb.value).filter(v => v);

    const newFriends = normalizeNames(document.getElementById('newFriends').value);

    if (present.length === 0 && !newFriends) {
        if (!confirm("目前出席人數為 0，確定要送出嗎？")) return;
    }

    showLoading("正在存入點名資料，請勿關閉網頁...");
    try {
        const res = await callAPI('submitAttendance', { groupName, date, present, absent, newFriends });
        if (res.success) {
            userNotification.success('點名成功！');
            document.querySelectorAll('.attendance-check').forEach(cb => cb.checked = false);
            document.getElementById('newFriends').value = '';
            if (groupCode) await loadGroupProgress();
        }
        else { userNotification.error('失敗：' + res.message); }
    } finally { hideLoading(); }
}
