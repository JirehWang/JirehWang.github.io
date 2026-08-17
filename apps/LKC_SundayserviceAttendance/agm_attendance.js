function agmMemberKey(member) {
  return String(member && (member.uid || member.memberUid || member.id || member.name) || '').trim();
}

function getAgmCategoryCounts(list) {
  const counts = { ALL: 0 };
  (Array.isArray(list) ? list : []).forEach(member => {
    counts.ALL++;
    const code = member && (member.categoryCode || member.category_code);
    if (code) counts[code] = (counts[code] || 0) + 1;
  });
  return counts;
}

function getAgmStateValue(state, member) {
  const key = agmMemberKey(member);
  return !!(state && (state[key] || state[member.name]));
}

function getAgmQuorumStats(list, checkedState, leaveState) {
  const members = Array.isArray(list) ? list : [];
  const state = checkedState || {};
  const leaves = leaveState || {};
  const cat1Members = members.filter(member => (member.categoryCode || member.category_code) === 'CAT_1');
  const leaveMembers = members.filter(member => getAgmStateValue(leaves, member));
  const cat1LeaveMembers = cat1Members.filter(member => getAgmStateValue(leaves, member));
  const effectiveCat1Members = cat1Members.filter(member => !getAgmStateValue(leaves, member));
  const effectiveTotal = cat1Members.length ? effectiveCat1Members.length : 204;
  const threshold = effectiveTotal > 0 ? Math.floor(effectiveTotal / 2) + 1 : 0;
  const checkedMembers = members.filter(member =>
    !getAgmStateValue(leaves, member) && getAgmStateValue(state, member)
  );
  const cat1Present = effectiveCat1Members.filter(member => getAgmStateValue(state, member)).length;

  return {
    checkedMembers,
    totalPresent: checkedMembers.length,
    cat1Present,
    effectiveTotal,
    threshold,
    presentCount: cat1Present,
    leaveCount: leaveMembers.length,
    cat1LeaveCount: cat1LeaveMembers.length,
    leaveMembers,
    isQuorumMet: effectiveTotal > 0 && cat1Present >= threshold
  };
}

function buildAgmAttendancePayload(list, checkedState, meetingTitle, scope, leaveState) {
  const members = Array.isArray(list) ? list : [];
  const state = checkedState || {};
  const stats = getAgmQuorumStats(members, state, leaveState);
  const checkedMembers = stats.checkedMembers;
  const checkedUids = [...new Set(checkedMembers
    .map(member => String(member.uid || member.memberUid || member.id || '').trim().toUpperCase())
    .filter(uid => /^LK\d+$/i.test(uid)))];
  const checkedNames = checkedMembers.map(member => String(member.name || '').trim()).filter(Boolean);
  const leaveUids = [...new Set(stats.leaveMembers
    .map(member => String(member.uid || member.memberUid || member.id || '').trim().toUpperCase())
    .filter(uid => /^LK\d+$/i.test(uid)))];
  const leaveNames = stats.leaveMembers.map(member => String(member.name || '').trim()).filter(Boolean);

  return {
    meetingTitle: meetingTitle || '會員大會點名紀錄',
    scope: scope || '',
    totalPresent: stats.totalPresent,
    cat1Present: stats.cat1Present,
    cat1Total: stats.effectiveTotal,
    cat1Leave: stats.cat1LeaveCount,
    quorumThreshold: stats.threshold,
    isQuorumMet: stats.isQuorumMet,
    checkedUids,
    checkedNames,
    leaveUids,
    leaveNames
  };
}

function normalizeAgmSessionRecord(session) {
  const item = session || {};
  const sessionId = String(item.sessionId || item.id || '').trim().toUpperCase();
  const meetingTitle = String(item.meetingTitle || item.sessionName || item.name || '').trim();
  return {
    sessionId,
    meetingTitle,
    sessionName: meetingTitle,
    status: String(item.status || 'OPEN').trim().toUpperCase() || 'OPEN',
    createdAt: String(item.createdAt || '').trim(),
    scope: sessionId ? 'AGM:' + sessionId : ''
  };
}

function buildAgmSessionQrUrl(sessionId, baseUrl) {
  const normalizedId = String(sessionId || '').trim().toUpperCase();
  if (!normalizedId) return '';
  const root = String(baseUrl || 'https://jirehwang.github.io/LKC1958_June_1.github.io/apps/LKC_SundayserviceAttendance/').replace(/[?].*$/, '').replace(/\/$/, '') + '/';
  return root + '?agmSession=' + encodeURIComponent(normalizedId) + '&agmRole=scanner';
}

function isAgmScannerQrEntry(entry) {
  const item = entry || {};
  return String(item.sessionId || '').trim() !== '' &&
    String(item.role || '').trim().toLowerCase() === 'scanner';
}

function isOfficialMemberActive(member) {
  if (!member || member.isActive === false) return false;
  const value = String(member.isActive == null ? '' : member.isActive).trim().toLowerCase();
  return !['false', '0', 'inactive', 'disabled', '停用', '否', 'n', 'no'].includes(value);
}

(function() {
  var activeCat = 'ALL';
  var checkedStateMap = {};
  var leaveStateMap = {};
  var agmSessions = [];
  var agmActiveSession = null;
  var agmMembers = [];
  var agmInitialized = false;
  var agmPollTimer = null;
  var agmLeaveMode = false;
  var agmQrRetryCount = 0;
  var agmScannerUserId = localStorage.getItem('agm_scanner_user_id');
  if (!agmScannerUserId) {
    agmScannerUserId = 'AGM_User_' + Math.floor(Math.random() * 1000000);
    localStorage.setItem('agm_scanner_user_id', agmScannerUserId);
  }

  function getAgmList() {
    return agmMembers;
  }

  function normalizeAgmMember(member) {
    return Object.assign({}, member, {
      name: String(member.name || '').trim(),
      uid: String(member.uid || member.memberUid || member.id || '').trim().toUpperCase(),
      categoryCode: member.categoryCode || member.category_code || '',
      categoryName: member.categoryName || member.category_name || '',
      isActive: isOfficialMemberActive(member)
    });
  }

  function getAgmStorageKey(prefix) {
    return prefix + '_' + (agmActiveSession ? agmActiveSession.sessionId : 'draft');
  }

  function loadAgmLocalState() {
    try {
      checkedStateMap = JSON.parse(localStorage.getItem(getAgmStorageKey('agm_checked_state')) || '{}');
      leaveStateMap = JSON.parse(localStorage.getItem(getAgmStorageKey('agm_leave_state')) || '{}');
    } catch (err) {
      checkedStateMap = {};
      leaveStateMap = {};
    }
  }

  function saveAgmLocalState() {
    localStorage.setItem(getAgmStorageKey('agm_checked_state'), JSON.stringify(checkedStateMap));
    localStorage.setItem(getAgmStorageKey('agm_leave_state'), JSON.stringify(leaveStateMap));
  }

  function hasAgmSession() {
    return !!(agmActiveSession && agmActiveSession.sessionId);
  }

  function isAgmScannerEntry() {
    return isAgmScannerQrEntry({
      sessionId: window.AGM_ENTRY_SESSION_ID,
      role: window.AGM_ENTRY_ROLE
    });
  }

  function applyAgmEntryMode() {
    const scannerEntry = isAgmScannerEntry();
    const sessionPanel = document.querySelector('.agm-session-panel');
    const titleInput = document.getElementById('agmMeetingTitle');

    if (sessionPanel) {
      sessionPanel.hidden = scannerEntry;
      sessionPanel.setAttribute('aria-hidden', String(scannerEntry));
    }
    if (titleInput) {
      const locked = scannerEntry || hasAgmSession();
      titleInput.readOnly = locked;
      titleInput.setAttribute('aria-readonly', String(locked));
    }
  }

  function getAgmSessionById(sessionId) {
    const normalized = String(sessionId || '').trim().toUpperCase();
    return agmSessions.find(session => session.sessionId === normalized) || null;
  }

  function renderAgmSessionOptions() {
    const select = document.getElementById('agmSessionSelect');
    if (!select) return;
    const currentId = agmActiveSession ? agmActiveSession.sessionId : '';
    select.innerHTML = '<option value="">請選擇已建立的場次</option>' + agmSessions.map(session =>
      '<option value="' + escapeHtml(session.sessionId) + '">' +
      escapeHtml(session.meetingTitle) + ' · ' + escapeHtml(session.createdAt || session.sessionId) +
      '</option>'
    ).join('');
    select.value = currentId;
  }

  function renderAgmSessionQr() {
    const canvas = document.getElementById('agmSessionQrCanvas');
    const empty = document.getElementById('agmSessionQrEmpty');
    const link = document.getElementById('agmSessionQrLink');
    const sessionUrl = hasAgmSession() ? buildAgmSessionQrUrl(agmActiveSession.sessionId, window.location.href) : '';
    if (canvas) canvas.hidden = !sessionUrl;
    if (canvas && sessionUrl && typeof QRious === 'undefined') {
      if (agmQrRetryCount < 12) {
        agmQrRetryCount++;
        setTimeout(renderAgmSessionQr, 250);
      }
    } else if (canvas && sessionUrl && typeof QRious !== 'undefined') {
      agmQrRetryCount = 0;
      new QRious({ element: canvas, value: sessionUrl, size: 190, level: 'H' });
    } else if (canvas) {
      agmQrRetryCount = 0;
      const context = canvas.getContext('2d');
      context.clearRect(0, 0, canvas.width, canvas.height);
    }
    if (empty) empty.hidden = !!sessionUrl;
    if (link) {
      link.textContent = sessionUrl || '建立場次後顯示 QR Code';
      link.href = sessionUrl || '#';
    }
  }

  function setAgmActiveSession(session) {
    agmActiveSession = session ? normalizeAgmSessionRecord(session) : null;
    if (agmActiveSession && !agmActiveSession.sessionId) agmActiveSession = null;
    if (agmActiveSession) {
      localStorage.setItem('agm_active_session_id', agmActiveSession.sessionId);
      loadAgmLocalState();
    } else {
      localStorage.removeItem('agm_active_session_id');
      checkedStateMap = {};
      leaveStateMap = {};
    }
    const titleInput = document.getElementById('agmMeetingTitle');
    const activeLabel = document.getElementById('agmActiveSessionLabel');
    const newBtn = document.getElementById('agmNewSessionBtn');
    if (titleInput) {
      if (agmActiveSession) titleInput.value = agmActiveSession.meetingTitle;
      titleInput.readOnly = !!agmActiveSession;
    }
    if (activeLabel) activeLabel.textContent = agmActiveSession
      ? '目前場次：' + agmActiveSession.meetingTitle
      : '尚未選擇場次；請先建立或選擇場次';
    if (newBtn) newBtn.disabled = !hasAgmSession() && !((titleInput && titleInput.value || '').trim());
    renderAgmSessionOptions();
    renderAgmSessionQr();
    updateAgmMemberCounts();
    if (agmInitialized) renderAgmGrid();
    applyAgmEntryMode();
    syncAgmDeviceMode();
  }

  function loadAgmSessions() {
    const requestedId = String(window.AGM_ENTRY_SESSION_ID || localStorage.getItem('agm_active_session_id') || '').trim().toUpperCase();
    const finish = function(list) {
      agmSessions = (Array.isArray(list) ? list : []).map(normalizeAgmSessionRecord).filter(session => session.sessionId && session.meetingTitle);
      renderAgmSessionOptions();
      const requested = getAgmSessionById(requestedId);
      if (requested) setAgmActiveSession(requested);
      else if (!hasAgmSession()) setAgmActiveSession(null);
    };
    if (typeof google === 'undefined' || !google.script || !google.script.run) {
      finish([]);
      return;
    }
    google.script.run
      .withSuccessHandler(finish)
      .withFailureHandler(function(err) {
        console.warn('會員大會場次載入失敗:', err);
        finish([]);
      })
      .getAgmSessions();
  }

  function setAgmMembers(list) {
    agmMembers = (Array.isArray(list) ? list : []).map(normalizeAgmMember).filter(member => member.name && member.isActive);
    agmMembers.forEach(member => {
      if (member.uid && checkedStateMap[member.name] && !checkedStateMap[member.uid]) {
        checkedStateMap[member.uid] = true;
      }
      if (member.uid && leaveStateMap[member.name] && !leaveStateMap[member.uid]) {
        leaveStateMap[member.uid] = true;
      }
    });
    saveAgmLocalState();
    updateAgmMemberCounts();
  }

  function updateAgmMemberCounts() {
    const counts = getAgmCategoryCounts(agmMembers);
    const labels = {
      ALL: '全部', CAT_1: '1. 應到', CAT_2: '2. 準會員', CAT_3: '3. 在外教派',
      CAT_4: '4. 外出', CAT_5: '5. 未陪餐', CAT_6: '6. 未陪餐籍在'
    };
    const pills = document.querySelectorAll('#agmCatPills button');
    pills.forEach(button => {
      const match = (button.getAttribute('onclick') || '').match(/filterAgmCat\('([^']+)'/);
      if (!match) return;
      const code = match[1];
      button.textContent = (labels[code] || code) + ' (' + (counts[code] || 0) + ')';
    });
    const quorumStats = getAgmQuorumStats(agmMembers, checkedStateMap, leaveStateMap);
    const cat1Total = quorumStats.effectiveTotal;
    const threshold = quorumStats.threshold;
    const totalEl = document.getElementById('agmCat1Total');
    const thresholdEl = document.getElementById('agmQuorumThreshold');
    const leaveEl = document.getElementById('agmCat1LeaveCount');
    if (totalEl) totalEl.textContent = cat1Total;
    if (thresholdEl) thresholdEl.textContent = threshold;
    if (leaveEl) leaveEl.textContent = quorumStats.cat1LeaveCount;
  }

  function getMeetingTitle() {
    return (agmActiveSession && agmActiveSession.meetingTitle) ||
      (document.getElementById('agmMeetingTitle')?.value || '').trim() || '會員大會點名紀錄';
  }

  function getAgmScope() {
    return agmActiveSession ? agmActiveSession.scope : '';
  }

  function isChecked(member) {
    return getAgmStateValue(checkedStateMap, member);
  }

  function isOnLeave(member) {
    return getAgmStateValue(leaveStateMap, member);
  }

  function escapeHtml(value) {
    return String(value || '').replace(/[&<>'"]/g, char => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
    }[char]));
  }

  function syncAgmDeviceMode() {
    if (!hasAgmSession() || typeof google === 'undefined' || !google.script || !google.script.run) return;
    google.script.run
      .withFailureHandler(() => {})
      .updateDeviceMode(agmScannerUserId, getAgmScope());
  }

  function syncAgmLeaveState(key, isOnLeave) {
    if (!hasAgmSession() || typeof google === 'undefined' || !google.script || !google.script.run) return;
    google.script.run
      .withFailureHandler(function(err) {
        console.warn('會員大會請假狀態同步失敗:', err);
      })
      .setAgmLeaveState(key, isOnLeave, getAgmScope(), agmScannerUserId);
  }

  window.selectAgmSession = function(sessionId) {
    if (isAgmScannerEntry()) return;
    const session = getAgmSessionById(sessionId);
    if (session) setAgmActiveSession(session);
  };

  window.startNewAgmSession = function() {
    if (isAgmScannerEntry()) return;
    setAgmActiveSession(null);
    const titleInput = document.getElementById('agmMeetingTitle');
    if (titleInput) {
      titleInput.readOnly = false;
      titleInput.focus();
      titleInput.select();
    }
  };

  window.createAgmSession = function() {
    if (isAgmScannerEntry()) return;
    const titleInput = document.getElementById('agmMeetingTitle');
    const title = String(titleInput && titleInput.value || '').trim();
    if (!title) {
      alert('請先輸入場次名稱');
      return;
    }
    if (typeof google === 'undefined' || !google.script || !google.script.run) {
      alert('目前無法連線後台，請稍後再試');
      return;
    }
    const button = document.getElementById('agmCreateSessionBtn');
    if (button) button.disabled = true;
    google.script.run
      .withSuccessHandler(function(session) {
        if (session) {
          agmSessions = [normalizeAgmSessionRecord(session)].concat(agmSessions.filter(item => item.sessionId !== session.sessionId));
          setAgmActiveSession(session);
          renderAgmSessionQr();
        }
        if (button) button.disabled = false;
      })
      .withFailureHandler(function(err) {
        if (button) button.disabled = false;
        alert('建立場次失敗：' + (err && err.message || err));
      })
      .createAgmSession(title, agmScannerUserId);
  };

  window.downloadAgmSessionQr = function() {
    if (!hasAgmSession()) {
      alert('請先建立或選擇場次');
      return;
    }
    renderAgmSessionQr();
    const canvas = document.getElementById('agmSessionQrCanvas');
    if (!canvas) return;
    const link = document.createElement('a');
    link.download = '會員大會場次QR_' + agmActiveSession.meetingTitle + '.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
  };

  window.openAgmViewer = function() {
    if (!hasAgmSession()) {
      alert('請先建立或選擇場次');
      return;
    }
    const viewerUrl = buildAgmSessionQrUrl(agmActiveSession.sessionId, window.location.href)
      .replace('agmRole=scanner', 'agmRole=viewer&agmViewer=1');
    const viewerWindow = window.open(viewerUrl, '_blank');
    if (!viewerWindow) alert('瀏覽器阻擋了觀看介面視窗，請允許彈出視窗後再試。');
  };

  function loadAgmMembers() {
    setAgmMembers([]);
    renderAgmGrid();
    syncAgmDeviceMode();
    loadAgmSessions();

    if (typeof google === 'undefined' || !google.script || !google.script.run) {
      startAgmCheckinPolling();
      return;
    }

    google.script.run
      .withSuccessHandler(function(list) {
        setAgmMembers(Array.isArray(list) ? list : []);
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
      const leaveClass = isOnLeave(member) ? 'on-leave' : '';
      const selectedClass = !leaveClass && isChecked(member) ? 'selected' : '';
      const catName = member.categoryName || member.category_name || '';
      const catCode = member.categoryCode || member.category_code || '';
      let catBadgeStyle = 'background: #e2e8f0; color: #334155;';
      if (catCode === 'CAT_1') catBadgeStyle = 'background: #dbeafe; color: #1e40af;';
      if (catCode === 'CAT_2') catBadgeStyle = 'background: #e0f2fe; color: #0369a1;';
      if (catCode === 'CAT_3') catBadgeStyle = 'background: #f1f5f9; color: #475569;';
      if (catCode === 'CAT_4') catBadgeStyle = 'background: #fef3c7; color: #92400e;';
      const stateTitle = leaveClass ? '已請假' : (selectedClass ? '已出席' : '未點名');
      return `<div class="agm-item ${selectedClass} ${leaveClass}" data-agm-key="${encodeURIComponent(key)}" title="${stateTitle}">
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

  window.toggleAgmLeaveMode = function() {
    agmLeaveMode = !agmLeaveMode;
    const button = document.getElementById('agmLeaveModeBtn');
    if (button) {
      button.classList.toggle('btn-warning', agmLeaveMode);
      button.classList.toggle('btn-outline-warning', !agmLeaveMode);
      button.textContent = agmLeaveMode ? '✅ 返回點名模式' : '📝 請假模式';
      button.setAttribute('aria-pressed', String(agmLeaveMode));
    }
  };

  window.toggleAgmItem = function(key) {
    if (!key) return;
    if (!hasAgmSession()) {
      alert('請先建立或選擇場次，才能開始點名');
      return;
    }
    const wasOnLeave = !!leaveStateMap[key];
    if (agmLeaveMode) {
      leaveStateMap[key] = !leaveStateMap[key];
      if (!leaveStateMap[key]) delete leaveStateMap[key];
      delete checkedStateMap[key];
    } else {
      checkedStateMap[key] = !checkedStateMap[key];
      if (!checkedStateMap[key]) delete checkedStateMap[key];
      delete leaveStateMap[key];
    }
    if (agmLeaveMode || wasOnLeave) syncAgmLeaveState(key, !!leaveStateMap[key]);
    saveAgmLocalState();
    renderAgmGrid();
  };

  function pollAgmCheckins() {
    if (!hasAgmSession() || typeof google === 'undefined' || !google.script || !google.script.run) return;
    const scope = getAgmScope();
    google.script.run
      .withSuccessHandler(function(result) {
        if (!result || result.scope !== scope) return;
        let changed = false;
        (result.leaveUids || []).forEach(uid => {
          const normalizedUid = String(uid || '').trim().toUpperCase();
          if (normalizedUid && !leaveStateMap[normalizedUid]) {
            leaveStateMap[normalizedUid] = true;
            delete checkedStateMap[normalizedUid];
            changed = true;
          }
        });
        (result.checkedUids || []).forEach(uid => {
          const normalizedUid = String(uid || '').trim().toUpperCase();
          if (normalizedUid && !leaveStateMap[normalizedUid] && !checkedStateMap[normalizedUid]) {
            checkedStateMap[normalizedUid] = true;
            changed = true;
          }
        });
        if (changed) {
          saveAgmLocalState();
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
    if (!hasAgmSession()) {
      alert('請先建立或選擇場次');
      return;
    }
    const scope = getAgmScope();
    checkedStateMap = {};
    leaveStateMap = {};
    saveAgmLocalState();
    renderAgmGrid();
    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run.withFailureHandler(function(err) {
        console.warn('清除會員大會 QR 暫存失敗：', err);
      }).clearAgmCheckinState(scope);
    }
  };

  window.toggleAgmScanner = function() {
    if (!hasAgmSession()) {
      alert('請先建立或選擇場次，再開啟會員 QR 掃描器');
      return;
    }
    const scope = getAgmScope();
    syncAgmDeviceMode();
    const scannerUrl = 'https://jirehwang.github.io/LKC1958_June_1.github.io/apps/qrcodescanner.github.io/?mode=' +
      encodeURIComponent(scope) + '&context=agm&operatorId=' + encodeURIComponent(agmScannerUserId || '') + '&userId=' + encodeURIComponent(agmScannerUserId || '');
    const scannerWindow = window.open(scannerUrl, '_blank');
    if (!scannerWindow) alert('⚠️ 瀏覽器阻擋了掃描器視窗，請允許彈出視窗後再試。');
  };

  window.submitAgmCheckins = async function() {
    if (!hasAgmSession()) {
      alert('請先建立或選擇場次');
      return;
    }
    const list = getAgmList();
    const meetingTitle = getMeetingTitle();
    const payload = buildAgmAttendancePayload(list, checkedStateMap, meetingTitle, getAgmScope(), leaveStateMap);
    payload.sessionId = agmActiveSession.sessionId;
    if (payload.totalPresent === 0) {
      alert('⚠️ 目前尚無任何會員點名簽到，無法送出紀錄！');
      return;
    }

    const quorumText = payload.isQuorumMet ? '✅ 已達超過50%成會門檻' : '⚠️ 未達超過50%成會門檻';
    const confirmMsg = `🏛️ 確認送出會員大會點名紀錄？\n\n` +
      `📌 會議名稱: ${meetingTitle}\n` +
      `👥 總簽到人數: ${payload.totalPresent} 人\n` +
      `🏛️ 應到會員出席: ${payload.cat1Present} / ${payload.cat1Total} 人 (${quorumText})\n\n` +
      `將送出點名紀錄至 Google Sheets「和會點名紀錄」工作表紀錄存檔。`;
    if (!confirm(confirmMsg)) return;

    if (typeof flushAttendanceTempQueue === 'function') {
      try {
        await flushAttendanceTempQueue();
        if (typeof flushAttendanceTempToBackendAsync === 'function') {
          await flushAttendanceTempToBackendAsync(payload.scope);
        }
      } catch (error) {
        alert('會員大會點名暫存尚未取得 Firebase ACK，請確認網路後重試。');
        return;
      }
    }

    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run
        .withSuccessHandler(function(res) {
          alert(res.message || '🎉 會員大會點名紀錄已成功送出並紀錄存檔！');
          checkedStateMap = {};
          leaveStateMap = {};
          saveAgmLocalState();
          renderAgmGrid();
        })
        .withFailureHandler(function(err) {
          alert('❌ 送出失敗：' + err.message);
        })
        .saveAgmAttendance(payload);
    } else {
      console.log('離線/模擬送出紀錄：', payload);
      alert(`🎉 [模擬] ${meetingTitle} 點名紀錄已成功儲存！\n總簽到: ${payload.totalPresent} 人 (${quorumText})`);
      leaveStateMap = {};
      saveAgmLocalState();
    }
  };

  function updateQuorumProgress() {
    const list = getAgmList();
    const quorumStats = getAgmQuorumStats(list, checkedStateMap, leaveStateMap);
    const totalCount = quorumStats.effectiveTotal;
    const threshold = quorumStats.threshold;
    const presentCount = quorumStats.presentCount;
    const percent = totalCount > 0 ? Math.min(100, Math.round((presentCount / totalCount) * 100)) : 0;
    const progressBar = document.getElementById('agmProgressBar');
    const presentEl = document.getElementById('agmPresentCount');
    const statusBadge = document.getElementById('agmQuorumStatusBadge');
    if (progressBar) progressBar.style.width = percent + '%';
    if (presentEl) presentEl.innerText = presentCount;
    if (statusBadge) {
      if (quorumStats.isQuorumMet) {
        statusBadge.className = 'badge bg-success text-white fw-bold px-2 py-1';
        statusBadge.innerText = '✅ 已達超過 50% 成會門檻 (' + percent + '%)';
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
    applyAgmEntryMode();
    loadAgmMembers();
  };
})();
