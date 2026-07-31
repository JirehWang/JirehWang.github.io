function agmMemberKey(member) {
  return String(member && (member.uid || member.memberUid || member.id || member.name) || '').trim();
}

function buildAgmAttendancePayload(list, checkedState, meetingTitle, scope) {
  const members = Array.isArray(list) ? list : [];
  const state = checkedState || {};
  const keyFor = member => String(member && (member.uid || member.memberUid || member.id || member.name) || '').trim();
  const checkedMembers = members.filter(member => {
    const key = keyFor(member);
    return !!(state[key] || state[member.name]);
  });
  const cat1Members = members.filter(member => (member.categoryCode || member.category_code) === 'CAT_1');
  const checkedUids = [...new Set(checkedMembers
    .map(member => String(member.uid || member.memberUid || member.id || '').trim().toUpperCase())
    .filter(uid => /^LK\d+$/i.test(uid)))];
  const checkedNames = checkedMembers.map(member => String(member.name || '').trim()).filter(Boolean);
  const cat1Present = cat1Members.filter(member => {
    const key = keyFor(member);
    return !!(state[key] || state[member.name]);
  }).length;
  const cat1Total = cat1Members.length || 204;

  return {
    meetingTitle: meetingTitle || '會員大會點名紀錄',
    scope: scope || '',
    totalPresent: checkedMembers.length,
    cat1Present,
    cat1Total,
    isQuorumMet: cat1Present >= Math.ceil(cat1Total * 0.5),
    checkedUids,
    checkedNames
  };
}

(function() {
  var activeCat = 'ALL';
  var checkedStateMap = JSON.parse(localStorage.getItem('agm_checked_state') || '{}');
  var agmMembers = [];
  var agmInitialized = false;
  var agmPollTimer = null;
  var agmScannerUserId = localStorage.getItem('agm_scanner_user_id');
  if (!agmScannerUserId) {
    agmScannerUserId = 'AGM_User_' + Math.floor(Math.random() * 1000000);
    localStorage.setItem('agm_scanner_user_id', agmScannerUserId);
  }

  function getAgmList() {
    return agmMembers.length ? agmMembers : (window.INITIAL_OFFICIAL_MEMBERS || []);
  }

  function normalizeAgmMember(member) {
    return Object.assign({}, member, {
      name: String(member.name || '').trim(),
      uid: String(member.uid || member.memberUid || member.id || '').trim().toUpperCase(),
      categoryCode: member.categoryCode || member.category_code || '',
      categoryName: member.categoryName || member.category_name || ''
    });
  }

  function setAgmMembers(list) {
    agmMembers = (Array.isArray(list) ? list : []).map(normalizeAgmMember).filter(member => member.name);
    agmMembers.forEach(member => {
      if (member.uid && checkedStateMap[member.name] && !checkedStateMap[member.uid]) {
        checkedStateMap[member.uid] = true;
      }
    });
    localStorage.setItem('agm_checked_state', JSON.stringify(checkedStateMap));
  }

  function getMeetingTitle() {
    return (document.getElementById('agmMeetingTitle')?.value || '').trim() || '會員大會點名紀錄';
  }

  function getAgmScope() {
    return 'AGM:' + getMeetingTitle().replace(/\s+/g, ' ').slice(0, 100);
  }

  function isChecked(member) {
    const key = agmMemberKey(member);
    return !!(checkedStateMap[key] || checkedStateMap[member.name]);
  }

  function escapeHtml(value) {
    return String(value || '').replace(/[&<>'"]/g, char => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
    }[char]));
  }

  function syncAgmDeviceMode() {
    if (typeof google === 'undefined' || !google.script || !google.script.run) return;
    google.script.run
      .withFailureHandler(() => {})
      .updateDeviceMode(agmScannerUserId, getAgmScope());
  }

  function loadAgmMembers() {
    setAgmMembers(window.INITIAL_OFFICIAL_MEMBERS || []);
    renderAgmGrid();
    syncAgmDeviceMode();

    if (typeof google === 'undefined' || !google.script || !google.script.run) {
      startAgmCheckinPolling();
      return;
    }

    google.script.run
      .withSuccessHandler(function(list) {
        if (Array.isArray(list) && list.length) setAgmMembers(list);
        renderAgmGrid();
        syncAgmDeviceMode();
        startAgmCheckinPolling();
      })
      .withFailureHandler(function(err) {
        console.warn('正式會員名單載入失敗，使用頁面內建名單：', err);
        startAgmCheckinPolling();
      })
      .getOfficialMembers();
  }

  window.filterAgmCat = function(catCode, btnEl) {
    activeCat = catCode;
    const container = document.getElementById('agmCatPills');
    if (container) {
      container.querySelectorAll('button').forEach(button => {
        button.className = 'btn btn-sm btn-outline-secondary fw-bold';
      });
      if (btnEl) btnEl.className = 'btn btn-sm btn-primary fw-bold active';
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
      filtered = filtered.filter(member => (member.categoryCode || member.category_code) === activeCat);
    }
    if (keyword) {
      filtered = filtered.filter(member => (member.name || '').toLowerCase().includes(keyword));
    }

    if (filtered.length === 0) {
      body.innerHTML = '<div class="text-center p-5 text-muted grid-column-full" style="grid-column: 1 / -1;">無符合條件的會員</div>';
      updateQuorumProgress();
      return;
    }

    body.innerHTML = filtered.map(member => {
      const key = agmMemberKey(member);
      const selectedClass = isChecked(member) ? 'selected' : '';
      const catName = member.categoryName || member.category_name || '';
      const catCode = member.categoryCode || member.category_code || '';
      let catBadgeStyle = 'background: #e2e8f0; color: #334155;';
      if (catCode === 'CAT_1') catBadgeStyle = 'background: #dbeafe; color: #1e40af;';
      if (catCode === 'CAT_2') catBadgeStyle = 'background: #e0f2fe; color: #0369a1;';
      if (catCode === 'CAT_3') catBadgeStyle = 'background: #f1f5f9; color: #475569;';
      if (catCode === 'CAT_4') catBadgeStyle = 'background: #fef3c7; color: #92400e;';
      return `<div class="agm-item ${selectedClass}" data-agm-key="${encodeURIComponent(key)}">
        <div class="agm-name">${escapeHtml(member.name)}</div>
        <span class="agm-cat-tag fw-bold" style="${catBadgeStyle}">${escapeHtml(catName)}</span>
      </div>`;
    }).join('');

    body.querySelectorAll('.agm-item').forEach(item => {
      item.addEventListener('click', function() {
        window.toggleAgmItem(decodeURIComponent(this.dataset.agmKey || ''));
      });
    });
    updateQuorumProgress();
  };

  window.toggleAgmItem = function(key) {
    if (!key) return;
    checkedStateMap[key] = !checkedStateMap[key];
    localStorage.setItem('agm_checked_state', JSON.stringify(checkedStateMap));
    renderAgmGrid();
  };

  function pollAgmCheckins() {
    if (typeof google === 'undefined' || !google.script || !google.script.run) return;
    const scope = getAgmScope();
    google.script.run
      .withSuccessHandler(function(result) {
        if (!result || result.scope !== scope) return;
        let changed = false;
        (result.checkedUids || []).forEach(uid => {
          const normalizedUid = String(uid || '').trim().toUpperCase();
          if (normalizedUid && !checkedStateMap[normalizedUid]) {
            checkedStateMap[normalizedUid] = true;
            changed = true;
          }
        });
        if (changed) {
          localStorage.setItem('agm_checked_state', JSON.stringify(checkedStateMap));
          renderAgmGrid();
        }
      })
      .withFailureHandler(function(err) {
        console.warn('會員大會 QR 同步失敗：', err);
      })
      .getAgmCheckinState(scope);
  }

  function startAgmCheckinPolling() {
    stopAgmCheckinPolling();
    pollAgmCheckins();
    agmPollTimer = setInterval(pollAgmCheckins, 5000);
  }

  window.stopAgmCheckinPolling = function() {
    if (agmPollTimer) clearInterval(agmPollTimer);
    agmPollTimer = null;
  };

  window.syncAgmCheckins = function() {
    renderAgmGrid();
    pollAgmCheckins();
  };

  window.resetAgmCheckins = function() {
    if (!confirm('⚠️ 是否確認重設/清空目前的會員大會簽到狀態？')) return;
    const scope = getAgmScope();
    checkedStateMap = {};
    localStorage.removeItem('agm_checked_state');
    renderAgmGrid();
    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run.withFailureHandler(function(err) {
        console.warn('清除會員大會 QR 暫存失敗：', err);
      }).clearAgmCheckinState(scope);
    }
  };

  window.toggleAgmScanner = function() {
    const scope = getAgmScope();
    syncAgmDeviceMode();
    const scannerUrl = 'https://jirehwang.github.io/LKC1958_June_1.github.io/apps/qrcodescanner.github.io/?userId=' +
      encodeURIComponent(agmScannerUserId) + '&mode=' + encodeURIComponent(scope) + '&context=agm';
    const scannerWindow = window.open(scannerUrl, '_blank');
    if (!scannerWindow) alert('⚠️ 瀏覽器阻擋了掃描器視窗，請允許彈出視窗後再試。');
  };

  window.submitAgmCheckins = function() {
    const list = getAgmList();
    const meetingTitle = getMeetingTitle();
    const payload = buildAgmAttendancePayload(list, checkedStateMap, meetingTitle, getAgmScope());
    if (payload.totalPresent === 0) {
      alert('⚠️ 目前尚無任何會員點名簽到，無法送出紀錄！');
      return;
    }

    const quorumText = payload.isQuorumMet ? '✅ 已達50%成會門檻' : '⚠️ 未達50%成會門檻';
    const confirmMsg = `🏛️ 確認送出會員大會點名紀錄？\n\n` +
      `📌 會議名稱: ${meetingTitle}\n` +
      `👥 總簽到人數: ${payload.totalPresent} 人\n` +
      `🏛️ 應到會員出席: ${payload.cat1Present} / ${payload.cat1Total} 人 (${quorumText})\n\n` +
      `將送出點名紀錄至 Google Sheets「和會點名紀錄」工作表紀錄存檔。`;
    if (!confirm(confirmMsg)) return;

    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run
        .withSuccessHandler(function(res) {
          alert(res.message || '🎉 會員大會點名紀錄已成功送出並紀錄存檔！');
          checkedStateMap = {};
          localStorage.removeItem('agm_checked_state');
          renderAgmGrid();
        })
        .withFailureHandler(function(err) {
          alert('❌ 送出失敗：' + err.message);
        })
        .saveAgmAttendance(payload);
    } else {
      console.log('離線/模擬送出紀錄：', payload);
      alert(`🎉 [模擬] ${meetingTitle} 點名紀錄已成功儲存！\n總簽到: ${payload.totalPresent} 人 (${quorumText})`);
    }
  };

  function updateQuorumProgress() {
    const list = getAgmList();
    const activeCommunicants = list.filter(member => (member.categoryCode || member.category_code) === 'CAT_1');
    const totalCount = activeCommunicants.length || 204;
    const threshold = Math.ceil(totalCount * 0.5);
    const presentCount = activeCommunicants.filter(isChecked).length;
    const percent = Math.min(100, Math.round((presentCount / totalCount) * 100));
    const progressBar = document.getElementById('agmProgressBar');
    const presentEl = document.getElementById('agmPresentCount');
    const statusBadge = document.getElementById('agmQuorumStatusBadge');
    if (progressBar) progressBar.style.width = percent + '%';
    if (presentEl) presentEl.innerText = presentCount;
    if (statusBadge) {
      if (presentCount >= threshold) {
        statusBadge.className = 'badge bg-success text-white fw-bold px-2 py-1';
        statusBadge.innerText = '✅ 已達 50% 成會門檻 (' + percent + '%)';
      } else {
        statusBadge.className = 'badge bg-warning text-dark fw-bold px-2 py-1';
        statusBadge.innerText = '⚠️ 尚差 ' + (threshold - presentCount) + ' 人成會 (' + percent + '%)';
      }
    }
  }

  window.initAgmAttendancePage = function() {
    if (agmInitialized) {
      renderAgmGrid();
      startAgmCheckinPolling();
      return;
    }
    agmInitialized = true;
    const titleInput = document.getElementById('agmMeetingTitle');
    if (titleInput) titleInput.addEventListener('input', syncAgmDeviceMode);
    loadAgmMembers();
  };
})();
