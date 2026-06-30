// showLoading / hideLoading / ensureAPIReady 由 config.js 提供。

// 🚀 網頁載入初始化邏輯 (包含專屬連結攔截)
window.onload = async () => {
    // 檢查中央路由是否就緒
    if (typeof window.churchAPI !== 'function') {
        userNotification.error("⚠️ 系統錯誤：安全路由 (config.js) 尚未載入！請聯絡管理員。");
        return;
    }

    const urlParams = new URLSearchParams(window.location.search);
    const queryId = urlParams.get('id');
    
    // 如果網址帶有 ?id=代碼
    if (queryId) {
        showLoading("正在驗證專屬連結...");
        try {
            // 🌟 使用中央路由發送請求
            const res = await window.churchAPI('findGroupByCode', { groupCode: queryId });
            
            if (res.success) {
                // 驗證成功，直接跳轉到該小組的點名頁面，並帶上名稱與代碼
                document.getElementById('overlay-text').innerText = "進入專屬小組中...";
                window.location.href = `group.html?name=${encodeURIComponent(res.groupName)}&code=${encodeURIComponent(res.encryptedCode)}`;
            } else {
                userNotification.warning("專屬連結無效或代碼錯誤！");
                hideLoading();
                fetchGroups(); // 若失敗，則載入正常首頁清單
            }
        } catch (e) {
            userNotification.error("驗證連結時發生錯誤。");
            hideLoading();
            fetchGroups();
        }
    } else {
        // 沒有帶 id，正常載入首頁的小組清單
        fetchGroups(); 
    }
};

// 載入小組按鈕清單
// 載入小組按鈕清單
async function fetchGroups() {
    showLoading("正在獲取最新小組清單...");
    const regListContainer = document.getElementById('regular-group-list');
    const happyListContainer = document.getElementById('happy-group-list');
    const inheritSelect = document.getElementById('inheritSourceGroup');
    try {
        // 🌟 使用中央路由發送請求
        const res = await window.churchAPI('getGroups');
        
        if (res.success) {
            regListContainer.innerHTML = '';
            happyListContainer.innerHTML = '';
            inheritSelect.innerHTML = '<option value="">-- 自訂新增同工 (不繼承) --</option>';

            const regularGroups = [];
            const activeHappyGroups = [];

            res.groups.forEach(group => {
                const type = group.type || '一般小組';
                const status = group.status || '顯示';

                if (type === '幸福小組') {
                    if (status === '顯示') {
                        activeHappyGroups.push(group);
                    }
                } else {
                    if (status !== '隱藏') {
                        regularGroups.push(group);
                    }
                }
            });

            // 1. 渲染常規小組
            regularGroups.forEach(group => {
                const btn = document.createElement('button');
                btn.className = 'tag-btn group-tag';
                btn.innerText = group.name;
                btn.onclick = () => enterGroup(group.name);
                regListContainer.appendChild(btn);

                // 填充繼承下拉選單
                const opt = document.createElement('option');
                opt.value = group.name;
                opt.innerText = group.name;
                inheritSelect.appendChild(opt);
            });

            // 常規小組的建立按鈕
            const createRegBtn = document.createElement('button');
            createRegBtn.className = 'tag-btn create-tag';
            createRegBtn.innerText = '➕ 創建新小組';
            createRegBtn.onclick = () => {
                toggleModal(true, '一般小組');
            };
            regListContainer.appendChild(createRegBtn);

            // 2. 渲染幸福小組
            activeHappyGroups.forEach(group => {
                const btn = document.createElement('button');
                btn.className = 'tag-btn group-tag happy-tag';
                btn.innerText = group.name;
                btn.onclick = () => enterGroup(group.name);
                happyListContainer.appendChild(btn);
            });

            // 幸福小組的建立按鈕
            const createHappyBtn = document.createElement('button');
            createHappyBtn.className = 'tag-btn create-tag';
            createHappyBtn.innerText = '➕ 創建新幸福小組';
            createHappyBtn.style.border = '2px dashed #E91E63';
            createHappyBtn.style.color = '#E91E63';
            createHappyBtn.style.background = '#fff0f3';
            createHappyBtn.onclick = () => {
                toggleModal(true, '幸福小組');
            };
            happyListContainer.appendChild(createHappyBtn);
        } else {
            regListContainer.innerHTML = `<p>讀取失敗：${res.message || '未知錯誤'}</p>`;
            happyListContainer.innerHTML = `<p>讀取失敗：${res.message || '未知錯誤'}</p>`;
        }
    } catch (e) {
        regListContainer.innerHTML = '<p>讀取失敗，請重新整理頁面</p>';
        happyListContainer.innerHTML = '<p>讀取失敗，請重新整理頁面</p>';
    } finally {
        hideLoading();
    }
}

function toggleModal(show, type = '一般小組') {
    const modal = document.getElementById('createModal');
    if (!modal) return;
    modal.style.display = show ? 'block' : 'none';
    if (show) {
        const title = document.getElementById('createModalTitle');
        const typeRadioReg = document.querySelector('input[name="newGroupType"][value="一般小組"]');
        const typeRadioHappy = document.querySelector('input[name="newGroupType"][value="幸福小組"]');
        
        if (type === '一般小組') {
            if (title) title.innerHTML = '⛪ 創建新常規小組';
            if (typeRadioReg) typeRadioReg.checked = true;
            toggleCreateInheritSection(false);
        } else {
            if (title) title.innerHTML = '🍀 創建新幸福小組';
            if (typeRadioHappy) typeRadioHappy.checked = true;
            toggleCreateInheritSection(true);
        }
        document.getElementById('newGroupName').value = '';
        document.getElementById('newGroupCode').value = '';
    }
}

// 建立新小組
async function createNewGroup() {
    const name = document.getElementById('newGroupName').value.trim();
    const code = document.getElementById('newGroupCode').value.trim();
    if (!name || !code) return userNotification.warning('請填寫完整資訊');

    const typeRadio = document.querySelector('input[name="newGroupType"]:checked');
    const type = typeRadio ? typeRadio.value : '一般小組';
    const associatedGroup = document.getElementById('inheritSourceGroup').value;

    showLoading("正在雲端建立小組並設定權限...");
    try {
        // 🌟 使用中央路由發送請求
        const res = await window.churchAPI('createGroup', { 
            groupName: name, 
            groupCode: code,
            groupType: type,
            associatedGroup: associatedGroup
        });

        (res.success ? userNotification.success : userNotification.warning)(res.message);
        if (res.success) {
            toggleModal(false);
            document.getElementById('newGroupName').value = '';
            document.getElementById('newGroupCode').value = '';
            document.getElementById('inheritSourceGroup').value = '';
            fetchGroups();
        }
    } catch (e) {
        userNotification.error("建立失敗，請稍後再試。");
    } finally {
        hideLoading();
    }
}

let currentVerifyingGroupName = "";

function toggleVerifyModal(show) {
    const modal = document.getElementById('verifyModal');
    if (!modal) return;
    modal.style.display = show ? 'block' : 'none';
    if (show) {
        document.getElementById('verifyGroupCode').value = '';
        document.getElementById('verifyError').style.display = 'none';
        setTimeout(() => {
            document.getElementById('verifyGroupCode').focus();
        }, 300);
    }
}

// 註冊 Enter 鍵解鎖
setTimeout(() => {
    const verifyCodeInput = document.getElementById('verifyGroupCode');
    if (verifyCodeInput) {
        verifyCodeInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                submitVerifyCode();
            }
        });
    }
}, 500);

// 手動點擊進入小組
async function enterGroup(groupName) {
    currentVerifyingGroupName = groupName;
    
    // 檢查 localStorage 快取 (具有 30 分鐘時效性)
    const cachedStr = localStorage.getItem(`group_code_${groupName}`);
    if (cachedStr) {
        try {
            const cached = JSON.parse(cachedStr);
            if (cached && cached.expiry && cached.expiry > Date.now()) {
                showLoading(`正在進入【${groupName}】...`);
                // 秒進：直接跳轉
                window.location.href = `group.html?name=${encodeURIComponent(groupName)}&code=${encodeURIComponent(cached.code)}`;
                return;
            }
        } catch (e) {
            // 解析失敗（例如舊版本直接存字串），直接視為過期無效並清理
        }
        localStorage.removeItem(`group_code_${groupName}`);
    }

    // 無快取，開啟自訂彈窗
    document.getElementById('verifyModalTitle').innerText = `🛡️ 驗證【${groupName}】身分`;
    toggleVerifyModal(true);
}

async function submitVerifyCode() {
    const code = document.getElementById('verifyGroupCode').value.trim();
    if (!code) return userNotification.warning('請輸入代碼');

    const errorEl = document.getElementById('verifyError');
    const submitBtn = document.getElementById('verifySubmitBtn');
    const codeInput = document.getElementById('verifyGroupCode');

    // 鎖定 UI
    errorEl.style.display = 'none';
    submitBtn.disabled = true;
    submitBtn.innerText = '⏳ 驗證中...';
    codeInput.disabled = true;

    try {
        const res = await window.churchAPI('verifyGroup', { groupName: currentVerifyingGroupName, groupCode: code });
        
        if (res.success) {
            // 寫入 localStorage 快取 (設定 30 分鐘過期時間)
            const cacheData = {
                code: res.encryptedCode,
                expiry: Date.now() + 30 * 60 * 1000 // 30 分鐘
            };
            localStorage.setItem(`group_code_${currentVerifyingGroupName}`, JSON.stringify(cacheData));
            
            toggleVerifyModal(false);
            showLoading("驗證成功，進入小組中...");
            window.location.href = `group.html?name=${encodeURIComponent(currentVerifyingGroupName)}&code=${encodeURIComponent(res.encryptedCode)}`;
        } else {
            errorEl.innerText = `❌ 驗證失敗：${res.message || '密碼錯誤'}`;
            errorEl.style.display = 'block';
            codeInput.disabled = false;
            codeInput.value = '';
            codeInput.focus();
            submitBtn.disabled = false;
            submitBtn.innerText = '進入小組';
        }
    } catch (e) {
        errorEl.innerText = '❌ 驗證時發生網路錯誤，請重試。';
        errorEl.style.display = 'block';
        codeInput.disabled = false;
        submitBtn.disabled = false;
        submitBtn.innerText = '進入小組';
    }
}

window.toggleVerifyModal = toggleVerifyModal;
window.submitVerifyCode = submitVerifyCode;
// --- 本週聚會人數彈窗 ---
async function openWeeklyReport() {
    const modal = document.getElementById('weeklyModal');
    const content = document.getElementById('weeklyReportContent');
    modal.style.display = 'flex';
    content.innerHTML = '<p style="color:#999; text-align:center; padding:2rem 0;">載入中...</p>';

    try {
        const res = await window.churchAPI('getWeeklyReport', {});
        if (!res.success) {
            content.innerHTML = `<p style="color:#e74c3c; text-align:center;">載入失敗：${res.message}</p>`;
            return;
        }
        if (res.data.length === 0) {
            content.innerHTML = `<p style="color:#999; text-align:center; padding:2rem 0;">本週尚無聚會紀錄</p>`;
            return;
        }

        const totalPeople = res.data.reduce((sum, g) => sum + g.total, 0);

        const rows = res.data.map((g, i) => {
            const newFriendBadge = g.newFriends > 0 
                ? `<span style="font-size:12px; background:#fff9c4; color:#f57f17; border-radius:99px; padding:2px 8px; margin-right:6px;">+${g.newFriends} 新朋友</span>` 
                : '';
            return `
                <div style="display:flex; align-items:center; gap:12px; padding:10px 0; border-bottom:0.5px solid #eee;">
                    <span style="font-size:13px; color:#aaa; width:20px; text-align:center;">${i + 1}</span>
                    <span style="font-size:14px; font-weight:500; flex:1;">${g.groupName}</span>
                    ${newFriendBadge}
                    <span style="font-size:13px; font-weight:500; background:#e8f5e9; color:#2e7d32; border-radius:99px; padding:2px 12px;">${g.total} 人</span>
                </div>
            `;
        }).join('');

        content.innerHTML = `
            <p style="font-size:12px; color:#999; margin:0 0 1rem;">
                📆 統計區間：${res.dateRange}
            </p>
            ${rows}
            <div style="display:flex; justify-content:space-between; margin-top:1rem; padding-top:0.75rem; border-top:1px solid #eee; font-size:13px; color:#666;">
                <span>共 ${res.data.length} 組有聚會紀錄</span>
                <span>本週合計 <strong>${totalPeople} 人</strong></span>
            </div>
        `;
    } catch (e) {
        content.innerHTML = `<p style="color:#e74c3c; text-align:center;">連線異常，請稍後再試</p>`;
    }
}

function closeWeeklyReport() {
    document.getElementById('weeklyModal').style.display = 'none';
}

// 點擊彈窗外部關閉 (延遲執行，確保 DOM 完全載入)
setTimeout(() => {
    const modal = document.getElementById('weeklyModal');
    if (modal) {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeWeeklyReport();
        });
    }
    const archModal = document.getElementById('archiveModal');
    if (archModal) {
        archModal.addEventListener('click', (e) => {
            if (e.target === archModal) toggleArchiveModal(false);
        });
    }
}, 500);

// ── 幸福小組頁籤與歷史封存功能 ──────────────────────────

function switchTab(tabType) {
    document.querySelectorAll('.tab-select-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content-section').forEach(sec => sec.style.display = 'none');
    
    if (tabType === 'regular') {
        document.getElementById('tab-regular').classList.add('active');
        document.getElementById('regular-groups-section').style.display = 'block';
    } else if (tabType === 'happy') {
        document.getElementById('tab-happy').classList.add('active');
        document.getElementById('happy-groups-section').style.display = 'block';
    }
}

function toggleCreateInheritSection(show) {
    document.getElementById('inheritSection').style.display = show ? 'block' : 'none';
    if (!show) {
        document.getElementById('inheritSourceGroup').value = '';
    }
}

function toggleArchiveModal(show) {
    document.getElementById('archiveModal').style.display = show ? 'block' : 'none';
}

async function openArchiveHistory() {
    toggleArchiveModal(true);
    const select = document.getElementById('archiveFileSelect');
    const detailContent = document.getElementById('archiveDetailContent');
    const loading = document.getElementById('archiveLoading');
    
    select.innerHTML = '<option value="">-- 請選擇 --</option>';
    detailContent.style.display = 'none';
    loading.style.display = 'block';

    try {
        const res = await window.churchAPI('happyGroup_getArchives', {});
        loading.style.display = 'none';
        
        if (res.success && res.files && res.files.length > 0) {
            res.files.forEach(file => {
                const opt = document.createElement('option');
                opt.value = file.id;
                opt.innerText = file.name.replace(".json", "");
                select.appendChild(opt);
            });
        } else if (res.success) {
            select.innerHTML = '<option value="">-- 目前無歷史封存檔案 --</option>';
        } else {
            userNotification.error("取得封存清單失敗：" + res.message);
        }
    } catch (e) {
        loading.style.display = 'none';
        userNotification.error("連線異常，取得封存清單失敗");
    }
}

async function loadSelectedArchiveContent() {
    const fileId = document.getElementById('archiveFileSelect').value;
    const detailContent = document.getElementById('archiveDetailContent');
    const loading = document.getElementById('archiveLoading');
    
    if (!fileId) {
        detailContent.style.display = 'none';
        return;
    }

    detailContent.style.display = 'none';
    loading.style.display = 'block';

    try {
        const res = await window.churchAPI('happyGroup_getArchiveContent', { fileId: fileId });
        loading.style.display = 'none';

        if (res.success && res.content) {
            const archive = res.content;
            
            // 1. 設定標題與日期
            document.getElementById('archiveGroupNameTitle').innerText = `🍀 ${archive.groupName}`;
            document.getElementById('archiveDateRange').innerText = `📅 聚會期程：${archive.startDate} ~ ${archive.endDate}`;

            // 2. 渲染成員名單 (Roster)
            const rosterDiv = document.getElementById('archiveRosterDiv');
            rosterDiv.innerHTML = '';
            
            if (archive.members && archive.members.length > 0) {
                archive.members.forEach(m => {
                    const badge = document.createElement('span');
                    const role = m.role || 'BEST';
                    let roleClass = 'role-best';
                    if (role === '福長' || role === '同工' || role === '核心同工') {
                        roleClass = 'role-core';
                    } else if (role === '一般同工') {
                        roleClass = 'role-general';
                    } else if (role === '陪伴同工') {
                        roleClass = 'role-companion';
                    }
                    
                    badge.className = `role-badge ${roleClass}`;
                    badge.style.width = 'auto';
                    badge.style.padding = '4px 10px';
                    badge.style.borderRadius = '20px';
                    badge.innerText = `${m.name} (${role})`;
                    rosterDiv.appendChild(badge);
                });
            } else {
                rosterDiv.innerHTML = '<p style="color: #999; margin: 0;">無成員資料</p>';
            }

            // 3. 渲染點名矩陣 (Attendance Matrix Table)
            const tableDiv = document.getElementById('archiveTableDiv');
            tableDiv.innerHTML = '';

            const records = archive.records || [];
            const members = archive.members || [];

            if (records.length === 0 || members.length === 0) {
                tableDiv.innerHTML = '<p style="color: #999; padding: 20px; text-align: center;">無點名紀錄資料</p>';
                detailContent.style.display = 'block';
                return;
            }

            // 建立表格
            const table = document.createElement('table');
            table.className = 'stats-dashboard';
            table.style.marginTop = '0';
            
            // 表頭: 姓名 | 性質 | [日期1] | [日期2] | ... | 出席次數/比例
            const thead = document.createElement('thead');
            const headerRow = document.createElement('tr');
            
            const thName = document.createElement('th');
            thName.innerText = '姓名';
            headerRow.appendChild(thName);

            const thRole = document.createElement('th');
            thRole.innerText = '身分';
            headerRow.appendChild(thRole);

            records.forEach(rec => {
                const thDate = document.createElement('th');
                const dParts = rec.date.split('-');
                thDate.innerText = dParts.length === 3 ? `${dParts[1]}/${dParts[2]}` : rec.date;
                thDate.title = rec.date;
                headerRow.appendChild(thDate);
            });

            const thRate = document.createElement('th');
            thRate.innerText = '出席率';
            headerRow.appendChild(thRate);
            
            thead.appendChild(headerRow);
            table.appendChild(thead);

            // 表身: 每個成員一行
            const tbody = document.createElement('tbody');
            
            members.forEach(member => {
                const tr = document.createElement('tr');
                
                const tdName = document.createElement('td');
                tdName.innerText = member.name;
                tdName.style.fontWeight = 'bold';
                tr.appendChild(tdName);

                const tdRole = document.createElement('td');
                tdRole.innerText = member.role || 'BEST';
                tr.appendChild(tdRole);

                let presentCount = 0;
                records.forEach(rec => {
                    const tdCell = document.createElement('td');
                    const isHelper = member.uid && member.uid.toUpperCase().startsWith("LK");
                    let attended = false;
                    
                    if (isHelper) {
                        attended = rec.present.indexOf(member.uid) !== -1 || rec.present.indexOf(member.name) !== -1;
                    } else {
                        attended = rec.present.indexOf(member.name) !== -1;
                    }

                    if (attended) {
                        tdCell.innerHTML = '<span style="color: #2e7d32; font-weight: bold;">✔</span>';
                        tdCell.style.backgroundColor = '#e8f5e9';
                        presentCount++;
                    } else {
                        tdCell.innerHTML = '<span style="color: #c62828;">✘</span>';
                        tdCell.style.backgroundColor = '#ffebee';
                    }
                    tr.appendChild(tdCell);
                });

                const tdRate = document.createElement('td');
                const rate = records.length > 0 ? ((presentCount / records.length) * 100).toFixed(0) : 0;
                tdRate.innerHTML = `<strong>${presentCount}/${records.length}</strong> (${rate}%)`;
                tr.appendChild(tdRate);

                tbody.appendChild(tr);
            });

            // 統計列：每週出席人數（同工 + BEST）
            const trTotal = document.createElement('tr');
            trTotal.style.background = '#f5f5f5';
            trTotal.style.fontWeight = 'bold';

            const tdTotalLabel = document.createElement('td');
            tdTotalLabel.innerText = '出席人數合計';
            tdTotalLabel.colSpan = 2;
            trTotal.appendChild(tdTotalLabel);

            records.forEach(rec => {
                const tdTotalCell = document.createElement('td');
                tdTotalCell.innerText = `${rec.total} 人`;
                trTotal.appendChild(tdTotalCell);
            });

            const tdTotalBlank = document.createElement('td');
            tdTotalBlank.innerText = '-';
            trTotal.appendChild(tdTotalBlank);

            tbody.appendChild(trTotal);

            table.appendChild(tbody);
            tableDiv.appendChild(table);

            detailContent.style.display = 'block';
        } else {
            userNotification.error("讀取封存內容失敗：" + res.message);
        }
    } catch (e) {
        loading.style.display = 'none';
        userNotification.error("連線異常，讀取檔案失敗");
    }
}

// ============================================================
//  🏗️ 群組與分類建立邏輯 (收合式 UI 版)
// ============================================================

let setupCachedData = null; // 儲存 districts, clusters, groups 關係
let setupAuthInfo = null;   // 儲存 { code, isAdmin, groupName }
let pendingSetupType = null; // 儲存待處理的建立類型 ('district', 'cluster', 'group')

// 1. 開啟/關閉群組建立折疊區 (展開/收合往左展開的三個小按鈕)
function toggleSetupCollapse(show) {
    const subContainer = document.getElementById('setupSubButtons');
    const area = document.getElementById('setupCollapseArea');
    if (!subContainer) return;

    const isExpanded = subContainer.style.opacity === '1';

    if (show === false || isExpanded) {
        // 收合子按鈕列
        subContainer.style.opacity = '0';
        subContainer.style.maxWidth = '0';
        subContainer.style.marginRight = '0';
        // 隱藏下方的表單區域與子面版
        if (area) area.style.display = 'none';
        const panelD = document.getElementById('panel-setup-district');
        const panelC = document.getElementById('panel-setup-cluster');
        const panelG = document.getElementById('panel-setup-group');
        if (panelD) panelD.style.display = 'none';
        if (panelC) panelC.style.display = 'none';
        if (panelG) panelG.style.display = 'none';
    } else {
        // 展開子按鈕列
        subContainer.style.opacity = '1';
        subContainer.style.maxWidth = '500px'; // 確保足夠寬度
        subContainer.style.marginRight = '8px';
    }
}

function toggleSetupAuthModal(show) {
    const modal = document.getElementById('setupAuthModal');
    if (!modal) return;
    modal.style.display = show ? 'block' : 'none';
    if (show) {
        document.getElementById('setupAuthCode').value = '';
        document.getElementById('setupAuthError').style.display = 'none';
        setTimeout(() => {
            document.getElementById('setupAuthCode').focus();
        }, 300);
    }
}

// 註冊 Enter 鍵送出
setTimeout(() => {
    const input = document.getElementById('setupAuthCode');
    if (input) {
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') submitSetupAuthCode();
        });
    }
}, 500);

// 2. 身分驗證提交 (與後台登入一致，呼叫 findGroupByCode)
async function submitSetupAuthCode() {
    const code = document.getElementById('setupAuthCode').value.trim();
    if (!code) return userNotification.warning('請輸入代碼');

    const errorEl = document.getElementById('setupAuthError');
    const submitBtn = document.getElementById('setupAuthSubmitBtn');
    const codeInput = document.getElementById('setupAuthCode');

    errorEl.style.display = 'none';
    submitBtn.disabled = true;
    submitBtn.innerText = '⏳ 驗證中...';
    codeInput.disabled = true;

    try {
        // 呼叫與後台驗證一致的 API
        const res = await window.churchAPI('findGroupByCode', { groupCode: code });
        
        if (res.success) {
            // 驗證成功，檢查當前點擊的 pendingSetupType 所需權限
            if ((pendingSetupType === 'district' || pendingSetupType === 'group') && !res.isAdmin) {
                errorEl.innerText = `❌ 驗證失敗：權限不足或代碼錯誤`;
                errorEl.style.display = 'block';
                codeInput.disabled = false;
                submitBtn.disabled = false;
                submitBtn.innerText = '驗證身分';
                return;
            }

            // 通過權限，儲存快取
            setupAuthInfo = {
                code: code,
                isAdmin: res.isAdmin,
                groupName: res.groupName || null
            };
            sessionStorage.setItem('setup_auth', JSON.stringify(setupAuthInfo));

            toggleSetupAuthModal(false);
            
            // 載入資料並展開 pendingSetupType
            const targetType = pendingSetupType;
            pendingSetupType = null;

            await loadSetupHierarchyData();
            
            if (targetType) {
                const area = document.getElementById('setupCollapseArea');
                const targetPanel = document.getElementById(`panel-setup-${targetType}`);
                if (area && targetPanel) {
                    const panelD = document.getElementById('panel-setup-district');
                    const panelC = document.getElementById('panel-setup-cluster');
                    const panelG = document.getElementById('panel-setup-group');
                    if (panelD) panelD.style.display = 'none';
                    if (panelC) panelC.style.display = 'none';
                    if (panelG) panelG.style.display = 'none';

                    targetPanel.style.display = 'block';
                    area.style.display = 'block';
                }
            }
        } else {
            errorEl.innerText = `❌ 驗證失敗：權限不足或代碼錯誤`;
            errorEl.style.display = 'block';
        }
    } catch (e) {
        errorEl.innerText = '❌ 驗證時發生網路錯誤，請重試。';
        errorEl.style.display = 'block';
    } finally {
        codeInput.disabled = false;
        submitBtn.disabled = false;
        submitBtn.innerText = '驗證身分';
    }
}

// 3. 載入資料並渲染收合面板
async function loadSetupHierarchyData() {
    showLoading("正在獲取最新分區與小組資料...");
    try {
        const res = await window.churchAPI('getDistrictsAndClusters', { authCode: setupAuthInfo.code });
        if (res.success) {
            setupCachedData = res;
            initGroupSetupUI();
        } else {
            userNotification.error("載入分區資料失敗：" + res.message);
        }
    } catch (e) {
        userNotification.error("連線異常，載入分區資料失敗");
    } finally {
        hideLoading();
    }
}

// 根據權限初始化 UI 與選項
function initGroupSetupUI() {
    const isAdmin = setupAuthInfo.isAdmin;

    // 隱藏非管理員不該看到的內部區塊（例如小組群底下的歸屬牧區選擇）
    document.querySelectorAll('.admin-only-block').forEach(el => {
        el.style.display = isAdmin ? 'block' : 'none';
    });

    // === (A) 填充分區 checkbox 列表 (新增牧區用) ===
    const districtClusterDiv = document.getElementById('setupDistrictClustersList');
    districtClusterDiv.innerHTML = '';
    setupCachedData.clusters.forEach(c => {
        const distLabel = c.districtName ? ` <span style="color:#888; font-size:11px;">(目前歸屬: ${c.districtName})</span>` : '';
        districtClusterDiv.innerHTML += `
            <label style="display: flex; align-items: center; gap: 6px; margin-bottom: 6px; cursor: pointer; font-size: 14px;">
                <input type="checkbox" name="setupDistrictClusters" value="${c.name}">
                ${c.name}${distLabel}
            </label>
        `;
    });
    if (setupCachedData.clusters.length === 0) {
        districtClusterDiv.innerHTML = '<span style="color:#999; font-size:13px;">目前無任何小組群可供選取</span>';
    }

    // === (B) 填充牧區下拉選單與小組清單 (新增小組群用) ===
    const clusterDistrictSelect = document.getElementById('newClusterDistrict');
    clusterDistrictSelect.innerHTML = '<option value="">-- 請選擇牧區 (選填) --</option>';
    setupCachedData.districts.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d.name;
        opt.innerText = d.name;
        clusterDistrictSelect.appendChild(opt);
    });

    const clusterGroupDiv = document.getElementById('setupClusterGroupsList');
    clusterGroupDiv.innerHTML = '';

    // 小組長只能拉入「無小組群歸屬」的小組；小組長自己所屬的小組預設勾選且 disabled
    // 最高權限可以看到所有「無小組群歸屬」的小組
    setupCachedData.groups.forEach(g => {
        const isSelfGroup = !isAdmin && g.name === setupAuthInfo.groupName;
        const hasCluster = !!g.clusterUuid; // 已有小組群

        if (isSelfGroup) {
            clusterGroupDiv.innerHTML += `
                <label style="display: flex; align-items: center; gap: 6px; margin-bottom: 6px; cursor: pointer; font-size: 14px; font-weight: bold; color: #E65100;">
                    <input type="checkbox" name="setupClusterGroups" value="${g.uuid}" checked disabled>
                    ${g.name} <span style="color:#FF9800; font-size:11px;">(自己的小組)</span>
                </label>
            `;
        } else if (!hasCluster) {
            clusterGroupDiv.innerHTML += `
                <label style="display: flex; align-items: center; gap: 6px; margin-bottom: 6px; cursor: pointer; font-size: 14px;">
                    <input type="checkbox" name="setupClusterGroups" value="${g.uuid}">
                    ${g.name}
                </label>
            `;
        }
    });

    // === (C) 填充小組群下拉選單 (新增小組用) ===
    const groupClusterSelect = document.getElementById('newSetupGroupClusterSelect');
    groupClusterSelect.innerHTML = '<option value="">-- 不歸屬 --</option>';
    setupCachedData.clusters.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c.name;
        opt.innerText = c.name;
        groupClusterSelect.appendChild(opt);
    });

    // 重設輸入框
    document.getElementById('newDistrictName').value = '';
    document.getElementById('newClusterNameInput').value = '';
    document.getElementById('newSetupGroupName').value = '';
    document.getElementById('newSetupGroupCode').value = '';
    document.getElementById('newSetupGroupClusterName').value = '';
    document.querySelector('input[name="newSetupGroupClusterOpt"][value="existing"]').checked = true;
    toggleSetupGroupClusterOpt();
}

// 4. 手風琴摺疊展開控制 (每次點擊均進行身分及權限檢查)
async function toggleAccordion(type) {
    // 檢查是否驗證過
    const cachedAuth = sessionStorage.getItem('setup_auth');
    if (!cachedAuth) {
        pendingSetupType = type;
        toggleSetupAuthModal(true);
        return;
    }

    try {
        setupAuthInfo = JSON.parse(cachedAuth);
    } catch (e) {
        sessionStorage.removeItem('setup_auth');
        pendingSetupType = type;
        toggleSetupAuthModal(true);
        return;
    }

    // 檢查權限是否足夠
    const isAdmin = setupAuthInfo.isAdmin;
    if ((type === 'district' || type === 'group') && !isAdmin) {
        userNotification.error("❌ 驗證失敗：權限不足或代碼錯誤");
        return;
    }

    const area = document.getElementById('setupCollapseArea');
    const targetPanel = document.getElementById(`panel-setup-${type}`);
    if (!area || !targetPanel) return;

    // 判斷當前面板是否已經顯示
    const isVisible = area.style.display === 'block' && targetPanel.style.display === 'block';

    if (isVisible) {
        // 如果已經顯示，再次點選則關閉整個折疊區
        area.style.display = 'none';
        targetPanel.style.display = 'none';
    } else {
        // 載入資料
        if (!setupCachedData) {
            await loadSetupHierarchyData();
        }
        
        // 隱藏其他面板，僅顯示目標面板
        const panelD = document.getElementById('panel-setup-district');
        const panelC = document.getElementById('panel-setup-cluster');
        const panelG = document.getElementById('panel-setup-group');
        if (panelD) panelD.style.display = 'none';
        if (panelC) panelC.style.display = 'none';
        if (panelG) panelG.style.display = 'none';

        targetPanel.style.display = 'block';
        area.style.display = 'block';
    }
}

function toggleSetupGroupClusterOpt() {
    const opt = document.querySelector('input[name="newSetupGroupClusterOpt"]:checked').value;
    const selectRow = document.getElementById('setup-group-cluster-select-row');
    const nameRow = document.getElementById('setup-group-cluster-name-row');
    if (opt === 'existing') {
        selectRow.style.display = 'block';
        nameRow.style.display = 'none';
    } else {
        selectRow.style.display = 'none';
        nameRow.style.display = 'block';
    }
}

// 5. 提交操作 API

// 🏰 新增牧區
async function submitDistrictSetup() {
    const name = document.getElementById('newDistrictName').value.trim();
    if (!name) return userNotification.warning('請輸入牧區名稱');

    const checkedBoxes = document.querySelectorAll('input[name="setupDistrictClusters"]:checked');
    const clusterUuids = Array.from(checkedBoxes).map(cb => cb.value); // 傳遞名稱列表

    showLoading('正在雲端建立牧區...');
    try {
        const res = await window.churchAPI('createDistrict', {
            name: name,
            clusterUuids: clusterUuids,
            authCode: setupAuthInfo.code
        });
        if (res.success) {
            userNotification.success('✅ 牧區建立成功！');
            toggleSetupCollapse(false);
            fetchGroups();
        } else {
            userNotification.warning(res.message || '建立失敗');
        }
    } catch (e) {
        userNotification.error('連線異常，建立牧區失敗');
    } finally {
        hideLoading();
    }
}

// 👥 新增小組群
async function submitClusterSetup() {
    const name = document.getElementById('newClusterNameInput').value.trim();
    if (!name) return userNotification.warning('請輸入小組群名稱');

    const districtUuid = setupAuthInfo.isAdmin ? document.getElementById('newClusterDistrict').value : '';

    const checkedBoxes = document.querySelectorAll('input[name="setupClusterGroups"]:checked');
    const groupUuids = Array.from(checkedBoxes).map(cb => cb.value);

    // 小組長建立時，把自己的小組加進去
    if (!setupAuthInfo.isAdmin) {
        const selfGroupObj = setupCachedData.groups.find(g => g.name === setupAuthInfo.groupName);
        if (selfGroupObj && !groupUuids.includes(selfGroupObj.uuid)) {
            groupUuids.push(selfGroupObj.uuid);
        }
    }

    showLoading('正在雲端建立小組群...');
    try {
        const res = await window.churchAPI('createGroupCluster', {
            name: name,
            districtUuid: districtUuid, // 牧區名稱
            groupUuids: groupUuids,
            authCode: setupAuthInfo.code
        });
        if (res.success) {
            userNotification.success('✅ 小組群建立成功！');
            toggleSetupCollapse(false);
            fetchGroups();
        } else {
            userNotification.warning(res.message || '建立失敗');
        }
    } catch (e) {
        userNotification.error('連線異常，建立小組群失敗');
    } finally {
        hideLoading();
    }
}

// ⛪ 新增小組
async function submitGroupSetup() {
    const name = document.getElementById('newSetupGroupName').value.trim();
    const code = document.getElementById('newSetupGroupCode').value.trim();
    if (!name || !code) return userNotification.warning('請填寫完整資訊');
    if (code.length < 4) return userNotification.warning('代碼至少需要 4 碼');

    const type = document.querySelector('input[name="newSetupGroupType"]:checked').value;
    const clusterOpt = document.querySelector('input[name="newSetupGroupClusterOpt"]:checked').value;

    let targetClusterUuid = '';
    let newClusterName = '';

    if (clusterOpt === 'existing') {
        targetClusterUuid = document.getElementById('newSetupGroupClusterSelect').value; // 小組群名稱
    } else {
        newClusterName = document.getElementById('newSetupGroupClusterName').value.trim();
        if (!newClusterName) return userNotification.warning('請輸入新小組群名稱');
    }

    showLoading('正在雲端建立小組並設定歸屬...');
    try {
        const res = await window.churchAPI('createGroup', {
            groupName: name,
            groupCode: code,
            groupType: type,
            targetClusterUuid: targetClusterUuid,
            newClusterName: newClusterName,
            authCode: setupAuthInfo.code
        });
        if (res.success) {
            userNotification.success('✅ 小組建立並歸屬成功！');
            toggleSetupCollapse(false);
            fetchGroups();
        } else {
            userNotification.warning(res.message || '建立失敗');
        }
    } catch (e) {
        userNotification.error('連線異常，建立小組失敗');
    } finally {
        hideLoading();
    }
}

// 暴露全域
window.toggleSetupCollapse = toggleSetupCollapse;
window.toggleSetupAuthModal = toggleSetupAuthModal;
window.submitSetupAuthCode = submitSetupAuthCode;
window.toggleAccordion = toggleAccordion;
window.toggleSetupGroupClusterOpt = toggleSetupGroupClusterOpt;
window.submitDistrictSetup = submitDistrictSetup;
window.submitClusterSetup = submitClusterSetup;
window.submitGroupSetup = submitGroupSetup;
