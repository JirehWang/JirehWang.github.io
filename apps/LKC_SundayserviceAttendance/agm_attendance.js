(function() {
  var activeCat = 'ALL';
  var checkedStateMap = JSON.parse(localStorage.getItem('agm_checked_state') || '{}');

  function getAgmList() {
    return window.INITIAL_OFFICIAL_MEMBERS || [];
  }

  window.filterAgmCat = function(catCode, btnEl) {
    activeCat = catCode;
    const container = document.getElementById('agmCatPills');
    if (container) {
      container.querySelectorAll('button').forEach(b => {
        b.className = "btn btn-sm btn-outline-secondary fw-bold";
      });
      if (btnEl) btnEl.className = "btn btn-sm btn-primary fw-bold active";
    }
    renderAgmGrid();
  };

  window.renderAgmGrid = function() {
    const body = document.getElementById('agmListBody');
    if (!body) return;

    const list = getAgmList();
    const keyword = (document.getElementById('agmSearchInput')?.value || '').trim().toLowerCase();

    let filtered = list;
    if (activeCat !== 'ALL') {
      filtered = filtered.filter(m => (m.category_code || m.categoryCode) === activeCat);
    }
    if (keyword) {
      filtered = filtered.filter(m => m.name.toLowerCase().includes(keyword));
    }

    if (filtered.length === 0) {
      body.innerHTML = '<div class="text-center p-5 text-muted grid-column-full" style="grid-column: 1 / -1;">無符合條件的會員</div>';
      updateQuorumProgress();
      return;
    }

    let html = '';
    filtered.forEach(m => {
      const isChecked = !!checkedStateMap[m.name];
      const selectedClass = isChecked ? 'selected' : '';
      const catName = m.category_name || m.categoryName || '';
      const catCode = m.category_code || m.categoryCode || '';

      let catBadgeStyle = 'background: #e2e8f0; color: #334155;';
      if (catCode === 'CAT_1') catBadgeStyle = 'background: #dbeafe; color: #1e40af;';
      if (catCode === 'CAT_2') catBadgeStyle = 'background: #e0f2fe; color: #0369a1;';
      if (catCode === 'CAT_3') catBadgeStyle = 'background: #f1f5f9; color: #475569;';
      if (catCode === 'CAT_4') catBadgeStyle = 'background: #fef3c7; color: #92400e;';

      html += `<div class="agm-item ${selectedClass}" onclick="toggleAgmItem('${m.name}')">
        <div class="agm-name">${m.name}</div>
        <span class="agm-cat-tag fw-bold" style="${catBadgeStyle}">${catName}</span>
      </div>`;
    });

    body.innerHTML = html;
    updateQuorumProgress();
  };

  window.toggleAgmItem = function(name) {
    checkedStateMap[name] = !checkedStateMap[name];
    localStorage.setItem('agm_checked_state', JSON.stringify(checkedStateMap));
    renderAgmGrid();
  };

  window.syncAgmCheckins = function() {
    renderAgmGrid();
  };

  window.resetAgmCheckins = function() {
    if (!confirm("⚠️ 是否確認重設/清空目前的會員大會簽到狀態？")) return;
    checkedStateMap = {};
    localStorage.removeItem('agm_checked_state');
    renderAgmGrid();
  };

  window.submitAgmCheckins = function() {
    const list = getAgmList();
    const meetingTitle = (document.getElementById('agmMeetingTitle')?.value || '').trim() || '會員大會點名紀錄';

    const checkedNames = Object.keys(checkedStateMap).filter(n => !!checkedStateMap[n]);
    if (checkedNames.length === 0) {
      alert("⚠️ 目前尚無任何會員點名簽到，無法送出紀錄！");
      return;
    }

    const cat1List = list.filter(m => (m.category_code || m.categoryCode) === 'CAT_1');
    const cat1Total = cat1List.length || 204;
    const cat1Threshold = Math.ceil(cat1Total * 0.5);

    let cat1Present = 0;
    cat1List.forEach(m => {
      if (checkedStateMap[m.name]) cat1Present++;
    });

    const isQuorumMet = cat1Present >= cat1Threshold;
    const quorumText = isQuorumMet ? "✅ 已達50%成會門檻" : "⚠️ 未達50%成會門檻";

    const confirmMsg = `🏛️ 確認送出會員大會點名紀錄？

` +
      `📌 會議名稱: ${meetingTitle}
` +
      `👥 總簽到人數: ${checkedNames.length} 人
` +
      `🏛️ 應到會員出席: ${cat1Present} / ${cat1Total} 人 (${quorumText})

` +
      `將送出點名紀錄至 Google Sheets「和會點名紀錄」工作表紀錄存檔。`;

    if (!confirm(confirmMsg)) return;

    const payload = {
      meetingTitle: meetingTitle,
      totalPresent: checkedNames.length,
      cat1Present: cat1Present,
      cat1Total: cat1Total,
      isQuorumMet: isQuorumMet,
      checkedNames: checkedNames
    };

    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run
        .withSuccessHandler(function(res) {
          alert(res.message || "🎉 會員大會點名紀錄已成功送出並紀錄存檔！");
        })
        .withFailureHandler(function(err) {
          alert("❌ 送出失敗：" + err.message);
        })
        .saveAgmAttendance(payload);
    } else {
      console.log("離線/模擬送出紀錄：", payload);
      alert(`🎉 [模擬] ${meetingTitle} 點名紀錄已成功儲存！
總簽到: ${checkedNames.length} 人 (${quorumText})`);
    }
  };

  function updateQuorumProgress() {
    const list = getAgmList();
    const activeCommunicants = list.filter(m => (m.category_code || m.categoryCode) === 'CAT_1');
    const totalCount = activeCommunicants.length || 204;
    const threshold = Math.ceil(totalCount * 0.5);

    let presentCount = 0;
    activeCommunicants.forEach(m => {
      if (checkedStateMap[m.name]) presentCount++;
    });

    const percent = Math.min(100, Math.round((presentCount / totalCount) * 100));

    const progressBar = document.getElementById('agmProgressBar');
    const presentEl = document.getElementById('agmPresentCount');
    const statusBadge = document.getElementById('agmQuorumStatusBadge');

    if (progressBar) progressBar.style.width = percent + '%';
    if (presentEl) presentEl.innerText = presentCount;

    if (statusBadge) {
      if (presentCount >= threshold) {
        statusBadge.className = "badge bg-success text-white fw-bold px-2 py-1";
        statusBadge.innerText = "✅ 已達 50% 成會門檻 (" + percent + "%)";
      } else {
        const needed = threshold - presentCount;
        statusBadge.className = "badge bg-warning text-dark fw-bold px-2 py-1";
        statusBadge.innerText = "⚠️ 尚差 " + needed + " 人成會 (" + percent + "%)";
      }
    }
  }

  // Initial render
  setTimeout(renderAgmGrid, 100);
})();
