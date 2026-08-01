function normalizeAgmViewerSession(session) {
  const item = session || {};
  const sessionId = String(item.sessionId || item.id || '').trim().toUpperCase();
  const meetingTitle = String(item.meetingTitle || item.sessionName || item.name || '').trim();
  return {
    sessionId,
    meetingTitle,
    scope: sessionId ? 'AGM:' + sessionId : '',
    status: String(item.status || 'OPEN').trim().toUpperCase() || 'OPEN',
    createdAt: String(item.createdAt || '').trim()
  };
}

function isAgmViewerMemberActive(member) {
  if (!member || member.isActive === false) return false;
  const value = String(member.isActive == null ? '' : member.isActive).trim().toLowerCase();
  return !['false', '0', 'inactive', 'disabled', '停用', '否', 'n', 'no'].includes(value);
}

(function() {
  let viewerSessions = [];
  let viewerMembers = [];
  let viewerActiveSession = null;
  let viewerCheckedUids = new Set();
  let viewerLeaveUids = new Set();
  let viewerPollTimer = null;
  let viewerInitialized = false;

  function hasViewerSession() {
    return !!(viewerActiveSession && viewerActiveSession.sessionId);
  }

  function viewerMemberKey(member) {
    return String(member && (member.uid || member.memberUid || member.id || member.name) || '').trim().toUpperCase();
  }

  function escapeViewerHtml(value) {
    return String(value || '').replace(/[&<>'"]/g, char => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
    }[char]));
  }

  function getViewerSession(sessionId) {
    const normalized = String(sessionId || '').trim().toUpperCase();
    return viewerSessions.find(session => session.sessionId === normalized) || null;
  }

  function getViewerStats() {
    const cat1 = viewerMembers.filter(member => (member.categoryCode || member.category_code) === 'CAT_1');
    const effectiveCat1 = cat1.filter(member => !viewerLeaveUids.has(viewerMemberKey(member)));
    const effectiveTotal = cat1.length ? effectiveCat1.length : 204;
    const threshold = effectiveTotal > 0 ? Math.floor(effectiveTotal / 2) + 1 : 0;
    const cat1Present = effectiveCat1.filter(member => viewerCheckedUids.has(viewerMemberKey(member))).length;
    return {
      effectiveTotal,
      threshold,
      cat1Present,
      cat1LeaveCount: cat1.length - effectiveCat1.length,
      leaveCount: viewerLeaveUids.size,
      isQuorumMet: effectiveTotal > 0 && cat1Present >= threshold
    };
  }

  function renderViewerSessions() {
    const select = document.getElementById('agmViewerSessionSelect');
    if (!select) return;
    select.innerHTML = '<option value="">請選擇要觀看的場次</option>' + viewerSessions.map(session =>
      '<option value="' + escapeViewerHtml(session.sessionId) + '">' +
      escapeViewerHtml(session.meetingTitle) + ' · ' + escapeViewerHtml(session.createdAt || session.sessionId) +
      '</option>'
    ).join('');
    select.value = viewerActiveSession ? viewerActiveSession.sessionId : '';
  }

  function renderViewerMembers() {
    const container = document.getElementById('agmViewerMemberList');
    if (!container) return;
    if (!hasViewerSession()) {
      container.innerHTML = '<div class="text-secondary small py-4 text-center" style="grid-column: 1 / -1;">選擇場次後顯示狀態</div>';
      return;
    }
    container.innerHTML = viewerMembers.map(member => {
      const key = viewerMemberKey(member);
      const isLeave = viewerLeaveUids.has(key);
      const isPresent = !isLeave && viewerCheckedUids.has(key);
      const status = isLeave ? '請假' : (isPresent ? '已到' : '未到');
      const klass = isLeave ? 'leave' : (isPresent ? 'present' : '');
      const category = member.categoryName || member.category_name || member.categoryCode || member.category_code || '';
      return '<div class="agm-viewer-member ' + klass + '">' +
        '<div class="agm-viewer-member-name">' + escapeViewerHtml(member.name) + '</div>' +
        '<div class="agm-viewer-member-status">' + escapeViewerHtml(category) + ' · ' + status + '</div>' +
        '</div>';
    }).join('');
  }

  function renderViewerStats(updatedAt) {
    const stats = getViewerStats();
    const percent = stats.effectiveTotal ? Math.min(100, Math.round(stats.cat1Present / stats.effectiveTotal * 100)) : 0;
    const setText = (id, value) => { const el = document.getElementById(id); if (el) el.textContent = value; };
    setText('agmViewerEligible', stats.effectiveTotal);
    setText('agmViewerLeave', stats.cat1LeaveCount);
    setText('agmViewerPresent', stats.cat1Present);
    setText('agmViewerThreshold', stats.threshold);
    setText('agmViewerUpdatedAt', updatedAt ? '最近同步：' + updatedAt : '尚未同步');
    const label = document.getElementById('agmViewerSessionLabel');
    if (label) label.textContent = hasViewerSession() ? '目前場次：' + viewerActiveSession.meetingTitle : '請選擇要觀看的場次';
    const bar = document.getElementById('agmViewerProgressBar');
    if (bar) bar.style.width = percent + '%';
    const badge = document.getElementById('agmViewerStatusBadge');
    if (badge) {
      badge.className = stats.isQuorumMet ? 'badge bg-success text-white' : 'badge bg-warning text-dark';
      badge.textContent = hasViewerSession() ? (stats.isQuorumMet ? '已達超過 50% 成會門檻' : '尚差 ' + Math.max(0, stats.threshold - stats.cat1Present) + ' 人') : '等待選擇場次';
    }
  }

  function applyViewerState(result) {
    if (!result || !hasViewerSession() || result.scope !== viewerActiveSession.scope) return;
    viewerCheckedUids = new Set((result.checkedUids || []).map(uid => String(uid || '').trim().toUpperCase()));
    viewerLeaveUids = new Set((result.leaveUids || []).map(uid => String(uid || '').trim().toUpperCase()));
    viewerLeaveUids.forEach(uid => viewerCheckedUids.delete(uid));
    renderViewerStats(new Date().toLocaleTimeString('zh-TW', { hour12: false }));
    renderViewerMembers();
  }

  function pollViewerState() {
    if (!hasViewerSession() || typeof google === 'undefined' || !google.script || !google.script.run) return;
    google.script.run
      .withSuccessHandler(applyViewerState)
      .withFailureHandler(function(err) { console.warn('會員大會觀看同步失敗：', err); })
      .getAgmCheckinState(viewerActiveSession.scope);
  }

  function startViewerPolling() {
    if (viewerPollTimer) clearInterval(viewerPollTimer);
    pollViewerState();
    viewerPollTimer = setInterval(pollViewerState, 5000);
  }

  function loadViewerMembers() {
    viewerMembers = [];
    const normalizeMember = member => Object.assign({}, member, {
      uid: String(member.uid || member.memberUid || member.id || '').trim().toUpperCase(),
      categoryCode: member.categoryCode || member.category_code || ''
    });
    if (typeof google === 'undefined' || !google.script || !google.script.run) {
      renderViewerMembers();
      renderViewerStats();
      return;
    }
    google.script.run
      .withSuccessHandler(function(list) {
        viewerMembers = (Array.isArray(list) ? list : [])
          .map(normalizeMember)
          .filter(member => member.name && isAgmViewerMemberActive(member));
        renderViewerMembers();
        renderViewerStats();
      })
      .withFailureHandler(function(err) { console.warn('會員大會觀看名單載入失敗：', err); })
      .getOfficialMembers();
  }

  function loadViewerSessions() {
    const requestedId = String(window.AGM_ENTRY_SESSION_ID || '').trim().toUpperCase();
    const finish = function(list) {
      viewerSessions = (Array.isArray(list) ? list : []).map(normalizeAgmViewerSession).filter(session => session.sessionId && session.meetingTitle);
      renderViewerSessions();
      const requested = getViewerSession(requestedId);
      if (requested) window.selectAgmViewerSession(requested.sessionId);
      else if (viewerSessions[0]) window.selectAgmViewerSession(viewerSessions[0].sessionId);
      else renderViewerStats();
    };
    if (typeof google === 'undefined' || !google.script || !google.script.run) {
      finish([]);
      return;
    }
    google.script.run
      .withSuccessHandler(finish)
      .withFailureHandler(function(err) { console.warn('會員大會觀看場次載入失敗：', err); finish([]); })
      .getAgmSessions();
  }

  window.selectAgmViewerSession = function(sessionId) {
    const session = getViewerSession(sessionId);
    viewerActiveSession = session;
    viewerCheckedUids = new Set();
    viewerLeaveUids = new Set();
    renderViewerSessions();
    renderViewerStats();
    renderViewerMembers();
    startViewerPolling();
  };

  window.refreshAgmViewer = function() {
    pollViewerState();
  };

  window.stopAgmViewerPolling = function() {
    if (viewerPollTimer) clearInterval(viewerPollTimer);
    viewerPollTimer = null;
  };

  window.initAgmViewerPage = function() {
    if (viewerInitialized) {
      renderViewerSessions();
      startViewerPolling();
      return;
    }
    viewerInitialized = true;
    loadViewerMembers();
    loadViewerSessions();
  };
})();
