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
async function fetchGroups() {
    showLoading("正在獲取最新小組清單...");
    const container = document.getElementById('group-list-container');
    try {
        // 🌟 使用中央路由發送請求
        const res = await window.churchAPI('getGroups');
        
        if (res.success) {
            container.innerHTML = '';
            res.groups.forEach(group => {
                const btn = document.createElement('button');
                btn.className = 'tag-btn group-tag';
                btn.innerText = group.name;
                btn.onclick = () => enterGroup(group.name);
                container.appendChild(btn);
            });

            const createBtn = document.createElement('button');
            createBtn.className = 'tag-btn create-tag';
            createBtn.innerText = '➕ 創建新小組';
            createBtn.onclick = () => toggleModal(true);
            container.appendChild(createBtn);
        } else {
            container.innerHTML = `<p>讀取失敗：${res.message || '未知錯誤'}</p>`;
        }
    } catch (e) {
        container.innerHTML = '<p>讀取失敗，請重新整理頁面</p>';
    } finally {
        hideLoading();
    }
}

function toggleModal(show) {
    document.getElementById('createModal').style.display = show ? 'block' : 'none';
}

// 建立新小組
async function createNewGroup() {
    const name = document.getElementById('newGroupName').value;
    const code = document.getElementById('newGroupCode').value;
    if (!name || !code) return userNotification.warning('請填寫完整資訊');

    showLoading("正在雲端建立小組並設定權限...");
    try {
        // 🌟 使用中央路由發送請求
        const res = await window.churchAPI('createGroup', { groupName: name, groupCode: code });

        (res.success ? userNotification.success : userNotification.warning)(res.message);
        if (res.success) {
            toggleModal(false);
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
    
    // 檢查 localStorage 快取
    const cachedCode = localStorage.getItem(`group_code_${groupName}`);
    if (cachedCode) {
        showLoading(`正在進入【${groupName}】...`);
        // 秒進：直接跳轉
        window.location.href = `group.html?name=${encodeURIComponent(groupName)}&code=${encodeURIComponent(cachedCode)}`;
        return;
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
            // 寫入 localStorage 快取
            localStorage.setItem(`group_code_${currentVerifyingGroupName}`, res.encryptedCode);
            
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
}, 500);
