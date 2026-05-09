  let currentStatsData = null; 
  let activeTab = '台語';
  let currentMode = 'range'; 
  let globalStatsGroupConfig = {};
  let currentStatsFetchId = 0;

  (function initStats() {
    const now = new Date();
    const toLocalISO = (d) => {
      const offset = d.getTimezoneOffset() * 60000;
      return new Date(d.getTime() - offset).toISOString().split('T')[0];
    };
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    document.getElementById('statsStart').value = toLocalISO(firstDay);
    document.getElementById('statsEnd').value = toLocalISO(now);
    document.getElementById('statsSingleDate').value = toLocalISO(now);
    loadStatsGroupConfig();
  })();

  function loadStatsGroupConfig() {
    google.script.run.withSuccessHandler(config => {
      globalStatsGroupConfig = config;
      renderStatsCategorySelect();
      updateStatsGroupSelect();
    }).getGroupConfig();
  }

  function renderStatsCategorySelect() {
    const catSelect = document.getElementById('statsCategorySelect');
    if (!catSelect) return;
    catSelect.innerHTML = '';
    for (let cat in globalStatsGroupConfig) {
      let opt = document.createElement('option');
      opt.value = cat;
      opt.text = (cat === '禮拜' ? '⛪ ' : (cat === '主日學' ? '📖 ' : '📂 ')) + cat;
      catSelect.add(opt);
    }
  }

  function updateStatsGroupSelect() {
    const catSelect = document.getElementById('statsCategorySelect');
    const grpSelect = document.getElementById('statsGroupSelect');
    if (!catSelect || !grpSelect) return;
    const currentCat = catSelect.value;
    grpSelect.innerHTML = ''; 
    if (currentCat && globalStatsGroupConfig[currentCat]) {
       let totalOpt = document.createElement('option');
       totalOpt.value = currentCat + '合計'; 
       totalOpt.text = `🌟 ${currentCat} 總合計`;
       grpSelect.add(totalOpt);
    }
    if (globalStatsGroupConfig[currentCat]) {
      globalStatsGroupConfig[currentCat].forEach(grp => {
        let cleanGrp = grp.replace('點名紀錄', '').trim(); 
        let opt = document.createElement('option');
        opt.value = cleanGrp;
        opt.text = cleanGrp;
        grpSelect.add(opt);
      });
    }
    fetchStatsData();
  }

  function toggleStatMode(mode) {
    currentMode = mode;
    if (mode === 'range') {
      document.getElementById('rangeInputArea').style.display = 'flex';
      document.getElementById('singleInputArea').style.display = 'none';
    } else {
      document.getElementById('rangeInputArea').style.display = 'none';
      document.getElementById('singleInputArea').style.display = 'flex';
    }
    document.getElementById('statsTableBody').innerHTML = '<tr><td colspan="4" class="text-center p-5 text-muted">模式已切換，請重新查詢</td></tr>';
    ['valMainStat', 'valMemberStat', 'valMemberMale', 'valMemberFemale', 'valNewFriendStat', 'valNfMale', 'valNfFemale'].forEach(id => {
       document.getElementById(id).innerText = '0';
    });
    const searchBox = document.getElementById('statsNameSearch');
    if (searchBox) searchBox.value = '';
    const trendRow = document.getElementById('trendAnalysisRow');
    if (trendRow) trendRow.style.display = 'none';
  }

  function getNextDayString(dateStr) {
    if (!dateStr) return "";
    var d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr; 
    d.setDate(d.getDate() + 1);
    var y = d.getFullYear();
    var m = ("0" + (d.getMonth() + 1)).slice(-2);
    var day = ("0" + d.getDate()).slice(-2);
    return y + "-" + m + "-" + day;
  }

  function fetchStatsData() {
    const catSelect = document.getElementById('statsCategorySelect');
    const grpSelect = document.getElementById('statsGroupSelect');
    if (!grpSelect || !grpSelect.value) return; 
    activeTab = grpSelect.value; 
    const currentCat = catSelect.value;
    const req = { 
      type: activeTab, 
      mode: currentMode,
      baseSheet: document.getElementById('baseSheetSource').value || "會友名單",
      targetGroups: []
    };
    if (activeTab.indexOf('合計') !== -1) {
      req.targetGroups = globalStatsGroupConfig[currentCat] || [];
    }
    if (currentMode === 'range') {
      req.start = document.getElementById('statsStart').value;
      const rawEnd = document.getElementById('statsEnd').value;
      if (!req.start || !rawEnd) return alert("請選擇起訖日期");
      req.end = getNextDayString(rawEnd); 
    } else {
      req.date = document.getElementById('statsSingleDate').value;
      if (!req.date) return alert("請選擇日期");
    }
    const tbody = document.getElementById('statsTableBody');
    tbody.innerHTML = '<tr><td colspan="4" class="text-center p-5"><div class="spinner-border text-primary mb-2"></div><div class="text-muted">正在分析數據...</div></td></tr>';
    const btn = document.querySelector('button[onclick="fetchStatsData()"]');
    if(btn) btn.disabled = true;
    currentStatsFetchId++;
    const myFetchId = currentStatsFetchId;
    google.script.run
      .withFailureHandler(function(err) {
          if (myFetchId !== currentStatsFetchId) return;
          tbody.innerHTML = `<tr><td colspan="4" class="text-center p-4 text-danger">錯誤<br><small>${err.message}</small></td></tr>`;
          if(btn) btn.disabled = false;
      })
      .withSuccessHandler(function(data) {
          if (myFetchId !== currentStatsFetchId) return;
          currentStatsData = data;
          renderStatsTable(data);
          if(btn) btn.disabled = false;
      })
      .getAttendanceStats(req);
  }

  function renderStatsTable(data) {
    if (!data) return;
    const nfMale = Number(data.nfMale || 0);
    const nfFemale = Number(data.nfFemale || 0);
    const presentMale = Number(data.presentMale || 0);
    const presentFemale = Number(data.presentFemale || 0);
    const presentCount = Number(data.presentCount || 0);
    const avgCount = Number(data.avgCount || 0);
    document.getElementById('valNewFriendStat').innerText = nfMale + nfFemale;
    document.getElementById('valNfMale').innerText = nfMale;
    document.getElementById('valNfFemale').innerText = nfFemale;
    document.getElementById('valMemberStat').innerText = presentCount;
    document.getElementById('valMemberMale').innerText = presentMale;
    document.getElementById('valMemberFemale').innerText = presentFemale;
    if (currentMode === 'range') {
      document.getElementById('lblMainStat').innerText = "平均總人數"; 
      document.getElementById('valMainStat').innerText = avgCount; 
    } else {
      document.getElementById('lblMainStat').innerText = "總出席人數";
      document.getElementById('valMainStat').innerText = presentCount + (nfMale + nfFemale);
    }
    const theadRow = document.getElementById('tableHeaderRow');
    if (currentMode === 'range') {
      theadRow.innerHTML = `<th class="ps-3">姓名</th><th class="text-center">性別</th><th class="text-center">次數</th><th style="width: 35%;">出席率</th>`;
    } else {
      theadRow.innerHTML = `<th class="ps-3">姓名</th><th class="text-center">性別</th><th class="text-center">狀態</th><th class="text-center">備註</th>`;
    }
    const tbody = document.getElementById('statsTableBody');
    let html = '';
    const thresholdInput = document.getElementById('threshold').value;
    const threshold = thresholdInput === "" ? 100 : parseInt(thresholdInput);
    if (data.details && data.details.length > 0) {
      data.details.forEach(row => {
        const gender = row.gender || '-';
        const genderColor = gender === '男' ? '#0d6efd' : (gender === '女' ? '#dc3545' : '#6c757d');
        const groupBadge = row.inGroup ? `<span class="badge bg-success ms-1 px-1 shadow-sm" style="font-size: 0.65em;">小組</span>` : '';
        if (currentMode === 'range') {
          if (row.rate <= threshold) {
            let barClass = 'bg-danger';
            if (row.rate > 80) barClass = 'bg-success';
            else if (row.rate > 50) barClass = 'bg-primary';
            else if (row.rate > 30) barClass = 'bg-warning';
            html += `<tr class="stat-row">
              <td class="ps-3 fw-bold stat-name">${row.name}${groupBadge}</td>
              <td class="text-center fw-bold" style="color:${genderColor}">${gender}</td>
              <td class="text-center fw-bold text-dark">${row.count}</td>
              <td class="pe-2">
                <div class="d-flex align-items-center">
                  <div class="progress flex-grow-1" style="height: 10px;">
                    <div class="progress-bar ${barClass}" style="width: ${row.rate}%"></div>
                  </div>
                  <span class="ms-2 small text-muted" style="width: 35px; text-align: right;">${row.rate}%</span>
                </div>
              </td>
            </tr>`;
          }
        } else {
          const statusBadge = row.attended ? 
            `<span class="badge bg-success bg-opacity-10 text-success border border-success px-3">已出席</span>` : 
            `<span class="badge bg-light text-muted border px-3">缺席</span>`;
          const nameStyle = row.attended ? "fw-bold text-dark" : "text-muted";
          html += `<tr class="stat-row">
            <td class="ps-3 ${nameStyle} stat-name">${row.name}${groupBadge}</td>
            <td class="text-center fw-bold" style="color:${genderColor}">${gender}</td>
            <td class="text-center">${statusBadge}</td>
            <td class="text-center small text-muted">-</td>
          </tr>`;
        }
      });
    } else {
      html = '<tr><td colspan="4" class="text-center p-4 text-muted">無資料</td></tr>';
    }
    tbody.innerHTML = html;
    filterStatsTable();
    const trendRow = document.getElementById('trendAnalysisRow');
    if (trendRow) trendRow.style.display = currentMode === 'range' ? '' : 'none';
  }

  function filterStatsTable() {
    const searchInput = document.getElementById('statsNameSearch');
    if (!searchInput) return;
    const kw = searchInput.value.trim().toLowerCase();
    const tbody = document.getElementById('statsTableBody');
    if (!tbody) return;
    tbody.querySelectorAll('tr.stat-row').forEach(row => {
      const nameCell = row.querySelector('.stat-name');
      if (nameCell) {
        const name = nameCell.innerText.trim().toLowerCase();
        row.style.display = (kw === "" || name.includes(kw)) ? '' : 'none';
      }
    });
  }

  function exportToCSV() {
    if (!currentStatsData || !currentStatsData.details) { alert("請先進行查詢"); return; }
    const baseSheetSelect = document.getElementById('baseSheetSource');
    const baseSheetName = baseSheetSelect ? baseSheetSelect.value : "會友名單";
    let csvRows = [];
    const nfMale = Number(currentStatsData.nfMale || 0);
    const nfFemale = Number(currentStatsData.nfFemale || 0);
    const presentMale = Number(currentStatsData.presentMale || 0);
    const presentFemale = Number(currentStatsData.presentFemale || 0);
    if (currentMode === 'range') {
        csvRows.push(['', `統計摘要 (${baseSheetName})`]);
        csvRows.push(['平均名單出席(男)', presentMale, '平均名單出席(女)', presentFemale]);
        csvRows.push(['期間新朋友(男)', nfMale, '期間新朋友(女)', nfFemale]);
        csvRows.push([]);
        csvRows.push(['姓名', '性別', '小組', '出席次數', '出席率(%)']);
        currentStatsData.details.forEach(row => {
            csvRows.push([row.name, row.gender, row.inGroup ? '是' : '否', row.count, row.rate + '%']);
        });
    } else {
        const date = document.getElementById('statsSingleDate').value;
        csvRows.push(['日期', date, '名單來源', baseSheetName]);
        csvRows.push(['名單出席(男)', presentMale, '名單出席(女)', presentFemale]);
        csvRows.push(['新朋友(男)', nfMale, '新朋友(女)', nfFemale]);
        csvRows.push([]); 
        csvRows.push(['姓名', '性別', '小組', '出席狀態']);
        currentStatsData.details.forEach(row => {
            csvRows.push([row.name, row.gender, row.inGroup ? '是' : '否', row.attended ? '出席' : '缺席']);
        });
    }
    const csvString = "\uFEFF" + csvRows.map(row => row.join(",")).join("\n");
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const fileName = currentMode === 'range' ? 
        `出席統計_區間_${activeTab}_${baseSheetName}.csv` : 
        `出席名單_${document.getElementById('statsSingleDate').value}_${activeTab}_${baseSheetName}.csv`;
    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // ===== 以下為出席頻率變化分析新功能 =====

  let trendRawData = null;

  function openTrendModal() {
    if (!currentStatsData) return alert("請先查詢");
    document.getElementById('trendModal').style.display = 'flex';
    runTrendAnalysis();
  }

  function closeTrendModal() {
    document.getElementById('trendModal').style.display = 'none';
  }

  document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closeTrendModal(); });

  function runTrendAnalysis() {
    const catSelect = document.getElementById('statsCategorySelect');
    const grpSelect = document.getElementById('statsGroupSelect');
    const start = document.getElementById('statsStart').value;
    const end = document.getElementById('statsEnd').value;
    if (!start || !end || !grpSelect || !grpSelect.value) return;

    document.getElementById('trendLoading').style.display = 'block';
    document.getElementById('trendResult').innerHTML = '';
    document.getElementById('trendSubtitle').innerHTML = '';

    const req = {
      type: grpSelect.value,
      start: start,
      end: getNextDayString(end),
      baseSheet: document.getElementById('baseSheetSource').value || '會友名單',
      targetGroups: grpSelect.value.indexOf('合計') !== -1
        ? (globalStatsGroupConfig[catSelect.value] || []) : []
    };

    google.script.run
      .withFailureHandler(function(err) {
        document.getElementById('trendLoading').style.display = 'none';
        document.getElementById('trendResult').innerHTML =
          `<div class="text-danger small p-3">錯誤：${err.message}</div>`;
      })
      .withSuccessHandler(function(data) {
        trendRawData = data;
        document.getElementById('trendLoading').style.display = 'none';
        applyTrendFilter();
      })
      .getAttendanceTrend(req);
  }

  function applyTrendFilter() {
    if (!trendRawData) return;
    const threshold = parseInt(document.getElementById('trendDropThreshold').value) || 20;
    renderTrendResult(trendRawData, threshold);
  }

function renderTrendResult(data, threshold) {
    const subtitle = document.getElementById('trendSubtitle');

    subtitle.innerHTML =
      `🔍 綜合衰退指數分析（歷史長度最多一年 vs 近期八週）<br>` +
      `<span class="text-muted small">歷史：<b>${data.periodHistory}</b>（${data.sessionsHistory} 場）</span><br>` +
      `<span class="text-muted small">近期：<b>${data.periodRecent}</b>（${data.sessionsRecent} 場）</span>`;

    const container = document.getElementById('trendResult');

    if (data.sessionsRecent === 0) {
      container.innerHTML = `<div class="text-center text-muted py-4 small">近期（八週內）無任何場次，無法進行分析。</div>`;
      return;
    }

    const dropped = (data.details || []).filter(function(r) { 
      return r.rateDrop >= threshold && r.dropScore > 0; 
    });

    // 依衰退指數排序
    dropped.sort(function(a, b) { return b.dropScore - a.dropScore; });

    let html = '';

    function buildSection(list) {
      if (list.length === 0) {
        return `<div class="text-muted small text-center py-2">無符合會友</div>`;
      }
      return list.map(function(r) {
        const gender = r.gender || '-';
        const gColor = gender === '男' ? '#0d6efd' : (gender === '女' ? '#dc3545' : '#888');
        const barA = Math.min(100, Math.round(r.historyRate));
        const barB = Math.min(100, Math.round(r.recentRate));
        const changePct = Math.round(r.rateDrop);
        const badgeColor = '#dc3545';
        const arrow = '↓';
        const cardBg = '#fffdf2';
        const warningHtml = r.missingThreeWeeks ? `<span class="text-danger fw-bold">${r.warningDesc}</span>` : '';

        return `
        <div style="border:1px solid #dee2e6; border-radius:8px; padding:10px 12px; margin-bottom:10px; background:${cardBg};">
          <div class="d-flex justify-content-between align-items-center mb-1">
            <span class="fw-bold" style="font-size:0.95rem;">${r.name}
              <span style="color:${gColor}; font-size:0.8rem; margin-left:4px;">${gender}</span>
            </span>
            <span class="badge" style="background:${badgeColor}; font-size:0.8rem;">${arrow} ${changePct}% (衰退指數: ${r.dropScore})</span>
          </div>
          <div class="d-flex align-items-center mb-1" style="gap:8px; font-size:0.78rem; color:#666;">
            <span style="width:26px;">歷史</span>
            <div style="flex:1; background:#e9ecef; border-radius:4px; height:8px;">
              <div style="width:${barA}%; background:#0d6efd; height:8px; border-radius:4px;"></div>
            </div>
            <span style="width:40px; text-align:right; color:#0d6efd; font-weight:600;">${r.historyRate}%</span>
          </div>
          <div class="d-flex align-items-center" style="gap:8px; font-size:0.78rem; color:#666;">
            <span style="width:26px;">近期</span>
            <div style="flex:1; background:#e9ecef; border-radius:4px; height:8px;">
              <div style="width:${barB}%; background:${badgeColor}; height:8px; border-radius:4px;"></div>
            </div>
            <span style="width:40px; text-align:right; font-weight:600; color:${badgeColor};">${r.recentRate}%</span>
          </div>
          <div style="font-size:0.75rem; color:#888; margin-top:5px; display:flex; justify-content:space-between; align-items:center;">
            <span>目前已連續缺席 <b>${r.consecutiveMisses}</b> 場</span>
            ${warningHtml}
          </div>
        </div>`;
      }).join('');
    }

    html += `
    <div class="mb-2">
      <div class="fw-bold mb-1" style="color:#dc3545; font-size:0.88rem;">
        📉 頻度突然降低（衰退幅度 ≥ ${threshold}%）　<span class="text-muted fw-normal">${dropped.length} 人</span>
      </div>
      ${buildSection(dropped)}
    </div>`;

    container.innerHTML = html;
  }

  function exportTrendCSV() {
    if (!trendRawData) { alert("請先進行分析"); return; }
    const threshold = parseInt(document.getElementById('trendDropThreshold').value) || 20;
    const data = trendRawData;
    const grpSelect = document.getElementById('statsGroupSelect');
    const groupName = grpSelect ? grpSelect.value : '';
    const start = document.getElementById('statsStart').value;
    const end = document.getElementById('statsEnd').value;

    let csvRows = [];
    csvRows.push(['出席頻率變化分析 (衰退指數排序)']);
    csvRows.push(['群組', groupName, '查詢區間', start + ' ~ ' + end]);
    csvRows.push(['歷史區間', data.periodHistory, '場次', data.sessionsHistory]);
    csvRows.push(['近期區間', data.periodRecent, '場次', data.sessionsRecent]);
    csvRows.push(['衰退幅度門檻', threshold + '%']);
    csvRows.push([]);
    csvRows.push(['類別', '姓名', '性別', '歷史出席率', '近期出席率', '衰退幅度', '連續缺席場數', '衰退指數', '備註']);

    const dropped = (data.details || []).filter(function(r) { return r.rateDrop >= threshold && r.dropScore > 0; });
    dropped.sort(function(a, b) { return b.dropScore - a.dropScore; });

    dropped.forEach(function(r) {
      csvRows.push(['出席減少', r.name, r.gender, r.historyRate + '%', r.recentRate + '%', r.rateDrop + '%', r.consecutiveMisses, r.dropScore, r.warningDesc || '']);
    });

    const csvString = "\uFEFF" + csvRows.map(function(row) { return row.join(','); }).join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `頻率變化分析_${groupName}_${start}_${end}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
