/**
 * app.js - LKC 奉獻管理系統 前端控制邏輯
 */

// 📅 初始化年份下拉選單
function initYearSelect(selectId) {
  const select = document.getElementById(selectId);
  if (!select) return;
  
  const currentYear = new Date().getFullYear();
  select.innerHTML = '';
  
  // 顯示當前年份及前兩年
  for (let y = currentYear; y >= currentYear - 2; y--) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.text = `${y} 年度`;
    select.add(opt);
  }
}

// 💵 格式化貨幣 (NT$ 1,234)
function formatCurrency(amount) {
  return 'NT$ ' + Number(amount).toLocaleString('zh-TW');
}

// ==========================================
// 1. 會友端查詢控制邏輯
// ==========================================
const MemberQueryController = {
  init() {
    initYearSelect('queryYear');
    
    const btnSearch = document.getElementById('btnSearch');
    if (btnSearch) {
      btnSearch.addEventListener('click', () => this.handleSearch());
    }
  },

  async handleSearch() {
    const codeInput = document.getElementById('memberCode');
    const nameInput = document.getElementById('memberName');
    const yearSelect = document.getElementById('queryYear');
    const errorAlert = document.getElementById('errorAlert');
    const resultSection = document.getElementById('resultSection');
    const spinner = document.getElementById('loadingSpinner');

    const code = (codeInput ? codeInput.value : '').trim();
    const name = (nameInput ? nameInput.value : '').trim();
    const year = yearSelect ? yearSelect.value : new Date().getFullYear().toString();

    // 清除先前狀態
    if (errorAlert) errorAlert.style.display = 'none';
    if (resultSection) resultSection.style.display = 'none';
    
    // 欄位防呆
    if (!code || !name) {
      this.showError('請完整輸入會友代號與姓名！');
      return;
    }

    // 啟動 Loading 動畫
    if (spinner) spinner.style.display = 'flex';

    try {
      const res = await OfferingAPI.queryMemberOffering(code, name, year);
      
      if (res && res.success) {
        this.renderResults(res);
      } else {
        this.showError(res.error || '查詢失敗，請重新確認代號或姓名。');
      }
    } catch (err) {
      this.showError('系統發生錯誤或連線逾時，請稍後再試。');
    } finally {
      if (spinner) spinner.style.display = 'none';
    }
  },

  showError(msg) {
    const errorAlert = document.getElementById('errorAlert');
    if (errorAlert) {
      errorAlert.textContent = '⚠️ ' + msg;
      errorAlert.style.display = 'block';
    }
  },

  renderResults(data) {
    const resultSection = document.getElementById('resultSection');
    const totalAmountEl = document.getElementById('totalAmount');
    const totalCountEl = document.getElementById('totalCount');
    const tableBody = document.querySelector('#offeringTable tbody');
    const memberTag = document.getElementById('memberTag');

    if (!resultSection) return;

    // 1. 隱私遮罩姓名展示
    if (memberTag) memberTag.textContent = `會友：${data.name}`;

    // 2. 累計計算
    let totalSum = 0;
    const records = data.records || [];
    
    tableBody.innerHTML = '';
    
    if (records.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="4" style="text-align:center; color:var(--text-muted);">本年度尚無此代號之奉獻紀錄</td></tr>`;
    } else {
      records.forEach(r => {
        totalSum += Number(r.amount) || 0;
        
        const tr = document.createElement('tr');
        // 依照類型給予不同樣式的 Badge
        let badgeClass = 'badge-primary';
        if (r.type === '感恩奉獻') badgeClass = 'badge-accent';
        if (r.type === '特別奉獻') badgeClass = 'badge-info';

        tr.innerHTML = `
          <td>${r.date}</td>
          <td><span class="badge ${badgeClass}">${r.type}</span></td>
          <td style="font-weight:700;">${formatCurrency(r.amount)}</td>
          <td style="color:var(--text-muted);">${r.notes || '-'}</td>
        `;
        tableBody.appendChild(tr);
      });
    }

    if (totalAmountEl) totalAmountEl.textContent = formatCurrency(totalSum);
    if (totalCountEl) totalCountEl.textContent = `${records.length} 筆`;

    // 顯示結果
    resultSection.style.display = 'block';
    resultSection.scrollIntoView({ behavior: 'smooth' });
  }
};

// ==========================================
// 2. 財務同工管理控制邏輯
// ==========================================
const AdminController = {
  parsedItems: [],

  init() {
    // 預設日期為今天
    const dateInput = document.getElementById('entryDate');
    if (dateInput) {
      const today = new Date();
      dateInput.value = window.formatYMD ? window.formatYMD(today) : today.toISOString().substring(0, 10);
    }

    const btnParse = document.getElementById('btnParse');
    if (btnParse) {
      btnParse.addEventListener('click', () => this.parseExcelData());
    }

    const btnUpload = document.getElementById('btnUpload');
    if (btnUpload) {
      btnUpload.addEventListener('click', () => this.handleUpload());
    }

    // 拖曳上傳預留
    this.initAiUploadZone();
  },

  // 解析從 Excel/Tab 分隔的複製貼上資料
  parseExcelData() {
    const textarea = document.getElementById('excelPasteArea');
    const tableBody = document.querySelector('#previewTable tbody');
    const previewSection = document.getElementById('previewSection');
    const statsInfo = document.getElementById('previewStats');

    const text = textarea ? textarea.value : '';
    if (!text.trim()) {
      alert('請先貼上資料！');
      return;
    }

    const lines = text.split('\n');
    this.parsedItems = [];
    tableBody.innerHTML = '';

    let totalAmount = 0;
    let validCount = 0;

    lines.forEach((line, index) => {
      if (!line.trim()) return;
      
      const columns = line.split('\t');
      if (columns.length < 2) return; // 至少要有代號和金額/項目

      const code = columns[0].trim();
      let type = '月定奉獻';
      let amount = 0;
      let notes = '';

      // 智慧型欄位判斷：
      // 如果欄位 2 是數字，說明沒有寫項目，默認為「月定奉獻」，欄位 2 即為金額
      const isCol2Num = !isNaN(Number(columns[1].replace(/[^0-9]/g, '')));
      
      if (isCol2Num) {
        amount = Number(columns[1].replace(/[^0-9]/g, ''));
        notes = columns.slice(2).join(' ').trim();
      } else {
        type = columns[1].trim();
        amount = columns[2] ? Number(columns[2].replace(/[^0-9]/g, '')) : 0;
        notes = columns.slice(3).join(' ').trim();
      }

      // 代號格式檢查 (例如 LKC001 或 TEST999)
      const isValidCode = code.length >= 3;
      const isValidAmount = amount > 0;
      const isRowValid = isValidCode && isValidAmount;

      if (isRowValid) {
        this.parsedItems.push({ code, type, amount, notes });
        totalAmount += amount;
        validCount++;
      }

      // 渲染預覽列 (紅字標記格式錯誤的列)
      const tr = document.createElement('tr');
      if (!isRowValid) {
        tr.style.backgroundColor = 'rgba(239, 68, 68, 0.08)';
        tr.style.color = '#dc2626';
      }

      tr.innerHTML = `
        <td>${isValidCode ? code : `⚠️ 格式錯誤 (${code})`}</td>
        <td><span class="badge badge-primary">${type}</span></td>
        <td style="font-weight:700;">${isValidAmount ? formatCurrency(amount) : '⚠️ 金額錯誤'}</td>
        <td>${notes || '-'}</td>
      `;
      tableBody.appendChild(tr);
    });

    if (statsInfo) {
      statsInfo.textContent = `成功解析：${validCount} 筆資料，合計金額：${formatCurrency(totalAmount)}`;
    }

    if (previewSection) {
      previewSection.style.display = 'block';
    }
  },

  // 財務數據上傳至 Google Sheet
  async handleUpload() {
    const adminTokenInput = document.getElementById('adminToken');
    const dateInput = document.getElementById('entryDate');
    const spinner = document.getElementById('uploadSpinner');

    const token = (adminTokenInput ? adminTokenInput.value : '').trim();
    const date = dateInput ? dateInput.value : '';

    if (!token) {
      alert('請先輸入管理權限密碼！');
      return;
    }
    
    // 簡易前端驗證
    if (token !== "LKC78213731") {
      alert('管理密碼不正確，拒絕上傳！');
      return;
    }

    if (this.parsedItems.length === 0) {
      alert('無有效的奉獻明細資料可上傳，請先解析貼上資料。');
      return;
    }

    const confirmUpload = confirm(`確定要將這 ${this.parsedItems.length} 筆資料上傳至 ${date} 奉獻明細中嗎？`);
    if (!confirmUpload) return;

    if (spinner) spinner.style.display = 'flex';

    try {
      const res = await OfferingAPI.adminAddOfferings(date, this.parsedItems, token);
      
      if (res && res.success) {
        alert(`✅ 登錄成功！共寫入 ${res.count} 筆奉獻明細。`);
        // 清空輸入
        const textarea = document.getElementById('excelPasteArea');
        if (textarea) textarea.value = '';
        const previewSection = document.getElementById('previewSection');
        if (previewSection) previewSection.style.display = 'none';
        this.parsedItems = [];
      } else {
        alert('❌ 寫入失敗：' + (res.error || '原因未知'));
      }
    } catch (e) {
      alert('❌ 系統發生錯誤：' + e.message);
    } finally {
      if (spinner) spinner.style.display = 'none';
    }
  },

  // AI 圖片辨識拖曳區初始化
  initAiUploadZone() {
    const zone = document.getElementById('aiUploadZone');
    const fileInput = document.getElementById('aiFileInput');
    const preview = document.getElementById('aiPreview');

    if (!zone || !fileInput) return;

    zone.addEventListener('click', () => fileInput.click());

    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      zone.style.borderColor = 'var(--primary)';
      zone.style.background = 'var(--primary-light)';
    });

    zone.addEventListener('dragleave', () => {
      zone.style.borderColor = 'var(--border)';
      zone.style.background = 'rgba(255, 255, 255, 0.3)';
    });

    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.style.borderColor = 'var(--border)';
      zone.style.background = 'rgba(255, 255, 255, 0.3)';
      
      if (e.dataTransfer.files.length > 0) {
        fileInput.files = e.dataTransfer.files;
        this.handleAiFileChange(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', () => {
      if (fileInput.files.length > 0) {
        this.handleAiFileChange(fileInput.files[0]);
      }
    });
  },

  handleAiFileChange(file) {
    const preview = document.getElementById('aiPreview');
    if (!preview) return;

    if (!file.type.startsWith('image/')) {
      alert('請選擇圖片檔案！');
      return;
    }

    const adminTokenInput = document.getElementById('adminToken');
    const token = (adminTokenInput ? adminTokenInput.value : '').trim();

    if (!token) {
      alert('請先在上面輸入財務同工權限密碼以進行 AI 辨識！');
      const fileInput = document.getElementById('aiFileInput');
      if (fileInput) fileInput.value = '';
      return;
    }
    if (token !== "LKC78213731") {
      alert('管理密碼不正確，拒絕進行 AI 辨識！');
      const fileInput = document.getElementById('aiFileInput');
      if (fileInput) fileInput.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = async (e) => {
      preview.src = e.target.result;
      preview.style.display = 'block';
      
      const dataUrl = e.target.result;
      const commaIdx = dataUrl.indexOf(',');
      if (commaIdx === -1) return;
      const base64Data = dataUrl.substring(commaIdx + 1);
      const mimeType = file.type;

      const spinner = document.getElementById('uploadSpinner');
      const previewSection = document.getElementById('previewSection');
      const statsInfo = document.getElementById('previewStats');
      const tableBody = document.querySelector('#previewTable tbody');

      if (previewSection) previewSection.style.display = 'block';
      if (spinner) spinner.style.display = 'flex';
      if (statsInfo) statsInfo.textContent = '🤖 AI 正在辨識圖片收據中，請稍候...';
      if (tableBody) tableBody.innerHTML = '';

      try {
        const res = await OfferingAPI.processReceiptImage(mimeType, base64Data, token);
        
        if (res && res.success && Array.isArray(res.items)) {
          this.parsedItems = [];
          tableBody.innerHTML = '';
          
          let totalAmount = 0;
          let validCount = 0;
          
          res.items.forEach(item => {
            const code = (item.code || '').trim();
            const name = (item.name || '').trim();
            const type = (item.type || '月定奉獻').trim();
            const amount = Number(item.amount) || 0;
            const notes = (item.notes || '').trim();
            
            const isValidCode = code.length >= 3;
            const isValidAmount = amount > 0;
            const isRowValid = isValidCode && isValidAmount;
            
            if (isRowValid) {
              this.parsedItems.push({ code, type, amount, notes });
              totalAmount += amount;
              validCount++;
            }
            
            const tr = document.createElement('tr');
            if (!isRowValid) {
              tr.style.backgroundColor = 'rgba(239, 68, 68, 0.08)';
              tr.style.color = '#dc2626';
            }
            
            const dispCode = name ? `${code} (${name})` : code;
            tr.innerHTML = `
              <td>${isValidCode ? dispCode : `⚠️ 格式錯誤 (${code || '無代號'})`}</td>
              <td><span class="badge badge-primary">${type}</span></td>
              <td style="font-weight:700;">${isValidAmount ? formatCurrency(amount) : '⚠️ 金額錯誤'}</td>
              <td>${notes || '-'}</td>
            `;
            tableBody.appendChild(tr);
          });
          
          if (statsInfo) {
            statsInfo.textContent = `🤖 AI 辨識完成！成功解析：${validCount} 筆資料，合計金額：${formatCurrency(totalAmount)}`;
          }
        } else {
          alert('❌ AI 辨識失敗：' + (res.error || '無法提取奉獻明細'));
          if (statsInfo) statsInfo.textContent = '❌ AI 辨識失敗';
        }
      } catch (err) {
        alert('❌ 系統發生錯誤：' + err.message);
        if (statsInfo) statsInfo.textContent = '❌ AI 辨識發生錯誤';
      } finally {
        if (spinner) spinner.style.display = 'none';
      }
    };
    reader.readAsDataURL(file);
  }
};

// ==========================================
// 3. 人名-編號查詢控制邏輯
// ==========================================
const SearchController = {
  password: "",

  init() {
    // 檢查 sessionStorage 中是否已有解鎖密碼
    const cachedPwd = sessionStorage.getItem('offering_search_pwd');
    if (cachedPwd === "LKC78213731") {
      this.password = cachedPwd;
      this.showSearchPanel();
    }

    const btnUnlock = document.getElementById('btnUnlock');
    if (btnUnlock) {
      btnUnlock.addEventListener('click', () => this.handleUnlock());
    }

    const btnSearchCode = document.getElementById('btnSearchCode');
    if (btnSearchCode) {
      btnSearchCode.addEventListener('click', () => this.handleSearch());
    }

    // 支援 Enter 鍵
    const searchPasswordInput = document.getElementById('searchPassword');
    if (searchPasswordInput) {
      searchPasswordInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') this.handleUnlock();
      });
    }

    const searchNameInput = document.getElementById('searchName');
    if (searchNameInput) {
      searchNameInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') this.handleSearch();
      });
    }
  },

  handleUnlock() {
    const pwdInput = document.getElementById('searchPassword');
    const errorAlert = document.getElementById('lockErrorAlert');
    const pwd = (pwdInput ? pwdInput.value : '').trim();

    if (errorAlert) errorAlert.style.display = 'none';

    if (!pwd) {
      this.showLockError('請輸入管理密碼！');
      return;
    }

    if (pwd !== "LKC78213731") {
      this.showLockError('密碼錯誤，請重新輸入！');
      return;
    }

    // 驗證成功
    this.password = pwd;
    sessionStorage.setItem('offering_search_pwd', pwd);
    this.showSearchPanel();
  },

  showLockError(msg) {
    const errorAlert = document.getElementById('lockErrorAlert');
    if (errorAlert) {
      errorAlert.textContent = '⚠️ ' + msg;
      errorAlert.style.display = 'block';
    }
  },

  showSearchPanel() {
    const lockScreen = document.getElementById('lockScreenSection');
    const searchPanel = document.getElementById('searchSection');
    if (lockScreen) lockScreen.style.display = 'none';
    if (searchPanel) searchPanel.style.display = 'block';
    
    // Focus search input
    const searchNameInput = document.getElementById('searchName');
    if (searchNameInput) searchNameInput.focus();
  },

  async handleSearch() {
    const nameInput = document.getElementById('searchName');
    const errorAlert = document.getElementById('searchErrorAlert');
    const resultSection = document.getElementById('resultSection');
    const spinner = document.getElementById('searchSpinner');
    const tableBody = document.querySelector('#resultTable tbody');
    const resultCount = document.getElementById('resultCount');

    const name = (nameInput ? nameInput.value : '').trim();

    if (errorAlert) errorAlert.style.display = 'none';
    if (resultSection) resultSection.style.display = 'none';

    if (!name) {
      this.showSearchError('請輸入姓名關鍵字！');
      return;
    }

    if (spinner) spinner.style.display = 'flex';

    try {
      const res = await OfferingAPI.searchMemberCode(name, this.password);
      
      if (res && res.success) {
        this.renderResults(res.results);
      } else {
        this.showSearchError(res.error || '查詢失敗，請重試。');
      }
    } catch (err) {
      this.showSearchError('系統發生錯誤，請聯絡同工。');
    } finally {
      if (spinner) spinner.style.display = 'none';
    }
  },

  showSearchError(msg) {
    const errorAlert = document.getElementById('searchErrorAlert');
    if (errorAlert) {
      errorAlert.textContent = '⚠️ ' + msg;
      errorAlert.style.display = 'block';
    }
  },

  renderResults(results) {
    const tableBody = document.querySelector('#resultTable tbody');
    const resultSection = document.getElementById('resultSection');
    const resultCount = document.getElementById('resultCount');

    if (!tableBody || !resultSection) return;

    tableBody.innerHTML = '';
    
    if (!results || results.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="2" style="text-align:center; color:var(--text-muted);">無匹配的會友資料</td></tr>`;
      if (resultCount) resultCount.textContent = '0 筆';
    } else {
      results.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td style="font-weight:700;">${r.name}</td>
          <td><span class="badge badge-primary" style="font-family:monospace; font-size:14px; letter-spacing:0.5px;">${r.code}</span></td>
        `;
        tableBody.appendChild(tr);
      });
      if (resultCount) resultCount.textContent = `${results.length} 筆`;
    }

    resultSection.style.display = 'block';
  }
};

// ==========================================
// 4. 入口路由判斷
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  const bodyId = document.body.id;
  
  if (bodyId === 'memberQueryPage') {
    MemberQueryController.init();
  } else if (bodyId === 'adminManagePage') {
    AdminController.init();
  } else if (bodyId === 'memberSearchPage') {
    SearchController.init();
  }
});
