let identifiedGroupName = "";
let isAdmin = false;
let debounceTimer;
let currentVerifyingCode = "";
let pendingVerificationPromise = null;
let nameDirectory = {};  // uid → name 反查表（從後端 RAW_MODE 回傳）
let verifiedCodeForQuery = ""; // 保存 URL 中已驗證的代碼

// UID → 姓名（找不到回傳原字串，相容新朋友純文字）
function resolveDisplayName(uidOrName) {
  if (!uidOrName) return "";
  const s = String(uidOrName).trim();
  if (/^LK\d+$/i.test(s)) return nameDirectory[s.toUpperCase()] || s;
  return s;
}

const splitRegex = /[^\u4e00-\u9fa5a-zA-Z0-9\s]+/; // 與後端一致的分隔符號

// showLoading / hideLoading / ensureAPIReady 由 config.js 提供。

// --- 系統啟動 ---
window.onload = async () => {
    try {
        showLoading("🚀 正在啟動系統通道...");
        await ensureAPIReady(); 
        
        // 新增：偵測網址是否帶有 id (加密 Token)
        const urlParams = new URLSearchParams(window.location.search);
        let queryId = urlParams.get('id');
        if (queryId) {
            const encryptGroupCode = window.encryptGroupCode || ((s) => s);
            const ENC_PREFIX = "enc_";
            if (queryId.indexOf(ENC_PREFIX) !== 0) {
                const encryptedId = encryptGroupCode(queryId);
                urlParams.set('id', encryptedId);
                const newUrl = window.location.pathname + '?' + urlParams.toString();
                window.history.replaceState({}, '', newUrl);
                queryId = encryptedId;
            }
            document.getElementById('groupCode').value = "******"; // 隱藏明文，顯示星號
            const res = await callAPI('findGroupByCode', { groupCode: queryId });
            if (res.success) {
                identifiedGroupName = res.groupName;
                isAdmin = res.isAdmin;
                const idRes = document.getElementById('idResult');
                idRes.className = 'status-badge status-ok';
                idRes.innerText = isAdmin ? '🛡️ 最高權限模式' : `✅ 小組：${res.groupName}`;
                verifiedCodeForQuery = queryId; // 記錄加密後的 Code 供查詢
                
                document.getElementById('adminGroupSelect').style.display = isAdmin ? 'inline-block' : 'none';
                const allMembersLabel = document.getElementById('typeAllMembersLabel');
                if (allMembersLabel) allMembersLabel.style.display = isAdmin ? 'inline-block' : 'none';
                if (isAdmin) await loadAdminOptions();
                
                // 自動執行數據載入
                await loadStats();
            } else {
                userNotification.warning("專屬連結無效或代碼錯誤！");
            }
        }
    } catch (e) {
        console.error(e);
        userNotification.error("系統啟動失敗：" + e.message);
    } finally {
        hideLoading();
    }
};

// --- API 呼叫 ---
async function callAPI(action, data = {}) {
    if (window.GroupSupabaseService && typeof window.GroupSupabaseService[action] === 'function') {
        try {
            const res = await window.GroupSupabaseService[action](data);
            if (res !== null && typeof res === 'object') return res;
        } catch (e) {
            console.warn('[GroupSupabase] Direct call error, fallback:', e);
        }
    }
    if (typeof window.churchAPI !== 'function') {
        userNotification.error("⚠️ 系統錯誤：安全路由尚未載入！");
        throw new Error("安全路由尚未載入");
    }
    return await window.churchAPI(action, data);
}

// --- 小組編號即時驗證 ---
async function verifyCode(code) {
    const idRes = document.getElementById('idResult');
    const adminSelect = document.getElementById('adminGroupSelect');
    
    if (code.length < 4) {
        idRes.innerText = '❌ 字數不足';
        idRes.className = 'status-badge status-err';
        identifiedGroupName = "";
        return;
    }
    
    currentVerifyingCode = code;
    // 加入轉圈圈 spinner 樣式與動畫
    idRes.innerHTML = '<span class="spinner-border spinner-border-sm me-1" role="status" style="width:11px; height:11px; display:inline-block; border: 2px solid currentColor; border-right-color: transparent; border-radius: 50%; animation: spin 0.75s linear infinite;"></span> 驗證中...';
    idRes.className = 'status-badge status-wait';

    try {
        const res = await callAPI('findGroupByCode', { groupCode: code });
        if (currentVerifyingCode !== code) return; // 避免競態條件

        if (res.success) {
            identifiedGroupName = res.groupName;
            isAdmin = res.isAdmin;
            idRes.className = 'status-badge status-ok';
            idRes.innerText = isAdmin ? '🛡️ 最高權限模式' : `✅ 小組：${res.groupName}`;
            adminSelect.style.display = isAdmin ? 'inline-block' : 'none';
            const allMembersLabel = document.getElementById('typeAllMembersLabel');
            if (allMembersLabel) allMembersLabel.style.display = isAdmin ? 'inline-block' : 'none';
            if (isAdmin) await loadAdminOptions();
        } else {
            identifiedGroupName = "";
            idRes.innerText = '❌ 查無此代碼';
            idRes.className = 'status-badge status-err';
            adminSelect.style.display = 'none';
        }
    } catch (err) {
        if (currentVerifyingCode === code) {
            idRes.innerText = '⚠️ 連線異常';
            idRes.className = 'status-badge status-err';
        }
    } finally {
        if (currentVerifyingCode === code) {
            pendingVerificationPromise = null;
        }
    }
}

// 註冊 CSS 旋轉動畫
if (typeof document !== 'undefined' && !document.getElementById('lkc-spin-style')) {
    const style = document.createElement('style');
    style.id = 'lkc-spin-style';
    style.textContent = '@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}';
    document.head && document.head.appendChild(style);
}

document.getElementById('groupCode').addEventListener('input', (e) => {
    const code = e.target.value.trim().toUpperCase();
    const idRes = document.getElementById('idResult');
    
    clearTimeout(debounceTimer);
    if (code.length === 0) {
        currentVerifyingCode = "";
        pendingVerificationPromise = null;
        identifiedGroupName = ""; // 當輸入框被清空時，清除已識別的小組名稱
        idRes.innerText = '等待輸入...';
        idRes.className = 'status-badge';
        return;
    }
    
    idRes.innerText = '等待中...';
    idRes.className = 'status-badge status-wait';

    // 縮短防抖至 400ms
    debounceTimer = setTimeout(() => {
        pendingVerificationPromise = verifyCode(code);
    }, 400);
});

async function loadAdminOptions() {
    const res = await callAPI('getGroups');
    const select = document.getElementById('adminGroupSelect');
    select.innerHTML = '<option value="ALL">-- 全小組彙整 --</option>';
    if (res.groups) {
        res.groups.forEach(g => {
            const opt = document.createElement('option');
            opt.value = g.name; opt.innerText = g.name;
            select.appendChild(opt);
        });
    }
}

// --- 數據查詢主入口 ---
async function loadStats() {
    const rawCodeInput = document.getElementById('groupCode');
    const rawCode = rawCodeInput ? rawCodeInput.value.trim().toUpperCase() : "";

    // 如果輸入了代碼但尚未開始驗證或驗證不同，立即取消防抖並執行驗證
    if (rawCode && rawCode !== "******" && rawCode !== currentVerifyingCode) {
        clearTimeout(debounceTimer);
        pendingVerificationPromise = verifyCode(rawCode);
    }

    // 如果有正在進行中的驗證，等待其完成
    if (pendingVerificationPromise) {
        showLoading("正在等待小組代碼驗證...");
        try {
            await pendingVerificationPromise;
        } catch (e) {
            console.error("驗證出錯:", e);
        } finally {
            hideLoading();
        }
    }

    if (!identifiedGroupName) return userNotification.warning('請先輸入正確的編號並等待識別');
    
    const reportType = document.querySelector('input[name="reportType"]:checked').value;
    const start = document.getElementById('startDate').value;
    const end = document.getElementById('endDate').value;
    const group = isAdmin ? document.getElementById('adminGroupSelect').value : identifiedGroupName;
    const code = (rawCode === "******" && verifiedCodeForQuery) ? verifiedCodeForQuery : (rawCode.startsWith("enc_") ? rawCode : rawCode.toUpperCase());

    showLoading("正在彙整報表數據...");
    
    try {
        if (reportType === "ALL_MEMBERS") {
            // 管理員專屬：總小組成員清單
            const res = await callAPI('getAllGroupMembers', { authCode: code });
            renderAllGroupMembers(res);
        } else if (reportType === "WEEKLY") {
            const targetGroup = (group === "ALL") ? "小組清單" : group;
            const res = await callAPI('getStats', {
                groupName: targetGroup,
                groupCode: code,
                startDate: "RAW_MODE"
            });
            if (res.nameDirectory) nameDirectory = res.nameDirectory;
            renderWeeklyStats(res, start, end);
        } else {
            const isAllGroups = (isAdmin && group === 'ALL');
            let res;
            if (isAllGroups) {
                res = await callAPI('getAllGroupsStats', { groupCode: code, startDate: start, endDate: end });
            } else {
                res = await callAPI('getStats', { groupName: group, groupCode: code, startDate: start, endDate: end });
            }
            renderMemberStats(res, start, end, isAllGroups);
        }
    } catch (e) {
        console.error(e);
        userNotification.error("查詢失敗：" + e.message);
    } finally {
        hideLoading();
    }
}

// --- 渲染：總小組成員清單（管理員專用）---
function renderAllGroupMembers(res) {
    if (!res.success) return userNotification.error(res.message || "讀取失敗");
    const thead = document.querySelector('#statsTable thead');
    const tbody = document.querySelector('#statsTable tbody');

    thead.innerHTML = `
        <tr>
            <th style="width: 18%;">姓名</th>
            <th style="width: 10%;">性別</th>
            <th style="width: 18%;">系統編號</th>
            <th style="width: 30%;">所屬小組</th>
            <th style="width: 24%;">身分</th>
        </tr>
    `;

    if (!res.data || res.data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:30px; color:#999;">沒有任何已歸組會友</td></tr>`;
        return;
    }

    // 依組別 → 姓名排序
    const sorted = res.data.slice().sort((a, b) => {
        const ga = (a.group || '').localeCompare(b.group || '');
        if (ga !== 0) return ga;
        return (a.name || '').localeCompare(b.name || '');
    });

    tbody.innerHTML = sorted.map(m => {
        const genderColor = m.gender === '男' ? '#0d6efd' : (m.gender === '女' ? '#dc3545' : '#6c757d');
        // 多組身分 → 拆 badge 顯示
        const roleHtml = _renderRoleBadges(m.role);
        // 多組所屬小組 → 拆顯示
        const groupHtml = (m.group || '').split(/[、,，]/).map(s => s.trim()).filter(s => s).map(g =>
            `<span style="background:#e3f2fd; color:#1565c0; padding:2px 8px; border-radius:4px; font-size:12px; margin:1px; display:inline-block;">${g}</span>`
        ).join(' ');
        return `
            <tr>
                <td style="font-weight:bold;">${m.name}</td>
                <td style="color:${genderColor}; font-weight:bold;">${m.gender || '-'}</td>
                <td style="font-family:monospace; color:#666;">${m.uid || '-'}</td>
                <td>${groupHtml || '<span style="color:#ccc;">-</span>'}</td>
                <td>${roleHtml || '<span style="color:#ccc;">-</span>'}</td>
            </tr>
        `;
    }).join('');
}

// 身分 badge 渲染（支援單組與多組「核心同工(A組)、一般同工(B組)」格式）
function _renderRoleBadges(roleStr) {
    if (!roleStr) return '';
    const COLORS = {
        '核心同工': 'background:#0d6efd; color:#fff;',
        '一般同工': 'background:#0dcaf0; color:#000;',
        '陪伴同工': 'background:#6c757d; color:#fff;',
        '小羊':     'background:#e9ecef; color:#333;'
    };
    return String(roleStr).split(/[、,，]/).map(s => s.trim()).filter(s => s).map(p => {
        const m = p.match(/^(.+?)\((.+?)\)$/);
        const r = m ? m[1].trim() : p;
        const g = m ? `<small style="opacity:0.8; margin-left:4px;">${m[2].trim()}</small>` : '';
        const style = COLORS[r] || COLORS['小羊'];
        return `<span style="${style} padding:2px 8px; border-radius:4px; font-size:12px; margin:1px; display:inline-block;">${r}${g}</span>`;
    }).join(' ');
}

// --- 渲染：每週出席人次 (包含出席組員與新朋友名單) ---
function renderWeeklyStats(res, start, end) {
    if (!res.success) return userNotification.error(res.message);
    const thead = document.querySelector('#statsTable thead');
    const tbody = document.querySelector('#statsTable tbody');

    thead.innerHTML = `
        <tr>
            <th style="width:15%">聚會日期</th>
            <th style="width:10%">出席人數</th>
            <th style="width:10%">新朋友</th>
            <th style="width:10%">總人次</th>
            <th style="text-align:left;">出席名單 (組員 / ✨新朋友)</th>
        </tr>
    `;

    const sLimit = start ? new Date(start).getTime() : 0;
    const eLimit = end ? new Date(end).getTime() : Infinity;

    const filteredRows = res.data.filter(row => {
        const d = new Date(row[0]).getTime();
        return d >= sLimit && d <= eLimit;
    });

    if (filteredRows.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5">此區間內查無點名紀錄</td></tr>`;
        return;
    }

    filteredRows.sort((a, b) => new Date(b[0]) - new Date(a[0]));

    tbody.innerHTML = filteredRows.map(row => {
        const dateStr = row[0] ? new Date(row[0]).toLocaleDateString() : "未知";

        // 解析出席者 UID → 姓名；新朋友維持純文字
        const presentRaw = row[1] ? row[1].toString().split(splitRegex).filter(n => n.trim()) : [];
        const presentArr = presentRaw.map(resolveDisplayName);
        const newFriendsArr = row[3] ? row[3].toString().split(splitRegex).filter(n => n.trim()) : [];

        const presentCount = presentArr.length;
        const newFriendsCount = newFriendsArr.length;
        const total = presentCount + newFriendsCount;

        const attendeeHTML = presentArr.map(name =>
            `<span style="display:inline-block; background:#e8f5e9; color:#2e7d32; padding:2px 8px; border-radius:4px; margin:2px; font-size:13px; border:1px solid #c8e6c9;">${name}</span>`
        ).join('');

        const newFriendsHTML = newFriendsArr.map(name =>
            `<span style="display:inline-block; background:#fff9c4; color:#f57f17; padding:2px 8px; border-radius:4px; margin:2px; font-size:13px; border:1px solid #ffe082;">✨ ${name}</span>`
        ).join('');

        return `
            <tr>
                <td style="font-weight:bold;">${dateStr}</td>
                <td style="color:#2ecc71; font-weight:bold; font-size:18px;">${presentCount}</td>
                <td style="color:#f1c40f; font-weight:bold; font-size:18px;">${newFriendsCount}</td>
                <td style="background:#f9f9f9; font-weight:bold; font-size:18px;">${total}</td>
                <td style="text-align:left; padding:10px;">
                    ${attendeeHTML}
                    ${newFriendsHTML}
                    ${(!attendeeHTML && !newFriendsHTML) ? '<span style="color:#ccc;">(無人出席)</span>' : ''}
                </td>
            </tr>
        `;
    }).join('');
}

// --- 渲染：組員出席率 ---
function renderMemberStats(res, start, end, showGroupCol) {
    if (!res.success) return userNotification.error(res.message);
    const thead = document.querySelector('#statsTable thead');
    const tbody = document.querySelector('#statsTable tbody');
    const isSingleDay = (start === end && start !== "");
    const showSunday = !showGroupCol;

    if (isSingleDay) {
        thead.innerHTML = `
            <tr>
                <th>姓名</th>
                ${showGroupCol ? '<th>所屬小組</th>' : ''}
                <th>小組出席</th>
                ${showSunday ? '<th>主日崇拜</th><th>主日學</th>' : ''}
            </tr>
        `;
        tbody.innerHTML = res.data.map(m => `
            <tr>
                <td style="font-weight:bold;">${m.name}</td>
                ${showGroupCol ? `<td><span class="badge">${m.group}</span></td>` : ''}
                <td style="font-size:20px;">${m.cell ? '✅' : '❌'}</td>
                ${showSunday ? `<td style="font-size:20px;">${m.sunday ? '✅' : '❌'}</td><td style="font-size:20px;">${m.school ? '✅' : '❌'}</td>` : ''}
            </tr>
        `).join('');
    } else {
        thead.innerHTML = `
            <tr>
                <th style="width:15%">姓名</th>
                ${showGroupCol ? '<th style="width:15%">所屬小組</th>' : ''}
                <th style="width: 23%;">🌱 小組出席率</th>
                ${showSunday ? '<th style="width: 23%;">⛪ 禮拜出席率</th><th style="width: 23%;">📖 主日學率</th>' : ''}
            </tr>
        `;
        tbody.innerHTML = res.data.map(m => `
            <tr>
                <td style="font-weight:bold;">${m.name}</td>
                ${showGroupCol ? `<td><span class="badge">${m.group}</span></td>` : ''}
                <td>${createProgressBar(m.cellStr, m.cellRate, 'color-cell')}</td>
                ${showSunday ? `<td>${createProgressBar(m.sundayStr, m.sundayRate, 'color-sunday')}</td><td>${createProgressBar(m.schoolStr, m.schoolRate, 'color-school')}</td>` : ''}
            </tr>
        `).join('');
    }
}

function createProgressBar(textStr, percentage, colorClass) {
    if (!textStr || textStr.endsWith("/0")) return `<span style="color:#aaa; font-size:12px;">無數據</span>`;
    const safePercentage = isNaN(percentage) ? 0 : parseFloat(percentage).toFixed(1);
    return `
        <div class="stat-box">
            <div class="stat-labels"><span>${textStr}</span><span>${safePercentage}%</span></div>
            <div class="prog-container"><div class="prog-bar ${colorClass}" style="width: ${safePercentage}%"></div></div>
        </div>
    `;
}

// --- Excel 匯出 ---
function exportToExcel() {
    const table = document.getElementById("statsTable");
    if (table.rows.length <= 1) return userNotification.warning('目前沒有資料可供匯出');
    showLoading("正在準備 Excel 檔案...");
    setTimeout(() => {
        let csv = "\ufeff";
        for (let i = 0; i < table.rows.length; i++) {
            const row = [], cols = table.rows[i].cells;
            for (let j = 0; j < cols.length; j++) row.push(cols[j].innerText.replace(/\n/g, ' '));
            csv += row.join(",") + "\r\n";
        }
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = `統計報表_${new Date().toLocaleDateString().replace(/\//g,'-')}.csv`;
        link.click();
        hideLoading();
    }, 500);
}
