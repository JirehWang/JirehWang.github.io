const state = {
  members: [],
  filters: { groups: [], ministries: [] },
  unresolved: [],
  selectedUid: '',
  participationExpanded: false
};

window.addEventListener('load', init);

async function init() {
  bindEvents();
  await loadMembers();
}

function bindEvents() {
  document.getElementById('refreshBtn').addEventListener('click', async () => {
    await refreshCaches();
  });
  document.getElementById('searchInput').addEventListener('input', renderMemberList);
  document.getElementById('groupFilter').addEventListener('change', renderMemberList);
  document.getElementById('ministryFilter').addEventListener('change', renderMemberList);
  document.getElementById('sundayFilter').addEventListener('change', renderMemberList);
  document.getElementById('sundaySchoolFilter').addEventListener('change', renderMemberList);
  document.getElementById('groupAttendanceFilter').addEventListener('change', renderMemberList);
  document.getElementById('participationFilter').addEventListener('change', renderMemberList);
  document.getElementById('participationToggle').addEventListener('click', () => {
    state.participationExpanded = !state.participationExpanded;
    renderMemberList();
  });
}

async function loadMembers() {
  showBusy('載入會友狀態...');
  try {
    await window.ensureAPIReady();
    const result = await window.churchAPI('getMembers', {});
    if (!result || !result.success) throw new Error(result && result.message ? result.message : '讀取失敗');
    state.members = result.members || [];
    state.filters = result.filters || { groups: [], ministries: [] };
    state.unresolved = result.unresolvedParticipants || [];
    populateFilters();
    renderMetrics();
    renderMemberList();
  } catch (error) {
    renderError(error);
  } finally {
    hideBusy();
  }
}

async function refreshCaches() {
  showBusy('刷新快取...');
  try {
    await window.churchAPI('refreshCaches', { forceRefresh: true });
    state.selectedUid = '';
    await loadMembers();
  } catch (error) {
    renderError(error);
    hideBusy();
  }
}

function populateFilters() {
  fillSelect('groupFilter', state.filters.groups || []);
  fillSelect('ministryFilter', state.filters.ministries || []);
  fillOptionSelect('sundayFilter', state.filters.attendanceOptions || []);
  fillOptionSelect('sundaySchoolFilter', state.filters.attendanceOptions || []);
  fillOptionSelect('groupAttendanceFilter', state.filters.attendanceOptions || []);
  fillOptionSelect('participationFilter', state.filters.participationOptions || []);
}

function fillSelect(id, values) {
  const select = document.getElementById(id);
  const current = select.value;
  select.innerHTML = '<option value="">全部</option>';
  values.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  if (values.includes(current)) select.value = current;
}

function fillOptionSelect(id, values) {
  const select = document.getElementById(id);
  const current = select.value;
  select.innerHTML = '';
  (values.length ? values : [{ value: '', label: '全部' }]).forEach(item => {
    const option = document.createElement('option');
    option.value = item.value || '';
    option.textContent = item.label || item.value || '全部';
    select.appendChild(option);
  });
  if (Array.from(select.options).some(option => option.value === current)) select.value = current;
}

function renderMetrics() {
  let servingGroupCount = 0;
  let servingMinistryCount = 0;
  
  state.members.forEach(m => {
    const groupCount = m.groupMinistries ? m.groupMinistries.length : 0;
    if (groupCount > 0) {
      servingGroupCount++;
    }
    const churchCount = m.churchMinistries ? m.churchMinistries.length : 0;
    const worshipCount = m.worshipPositions ? m.worshipPositions.length : (m.worship && m.worship.positions ? m.worship.positions.length : 0);
    if ((churchCount + worshipCount) > 0) {
      servingMinistryCount++;
    }
  });

  document.getElementById('metricMembers').textContent = state.members.length;
  document.getElementById('metricServingGroup').textContent = servingGroupCount;
  document.getElementById('metricServingMinistry').textContent = servingMinistryCount;
  document.getElementById('metricGroups').textContent = (state.filters.groups || []).length;
  document.getElementById('metricMinistries').textContent = (state.filters.ministries || []).length;
  document.getElementById('metricUnresolved').textContent = state.unresolved.length;
}

function renderMemberList() {
  const list = document.getElementById('memberList');
  const template = document.getElementById('memberRowTemplate');
  const query = document.getElementById('searchInput').value.trim().toLowerCase();
  const group = document.getElementById('groupFilter').value;
  const ministry = document.getElementById('ministryFilter').value;
  const sunday = document.getElementById('sundayFilter').value;
  const sundaySchool = document.getElementById('sundaySchoolFilter').value;
  const groupAttendance = document.getElementById('groupAttendanceFilter').value;
  const participation = document.getElementById('participationFilter').value;

  const rows = state.members.filter(member => {
    const matchesQuery = !query ||
      String(member.name || '').toLowerCase().includes(query) ||
      String(member.uid || '').toLowerCase().includes(query);
    const matchesGroup = !group || (member.groups || []).some(g => g.name === group);
    const matchesMinistry = !ministry ||
      (member.groupMinistries || []).includes(ministry) ||
      (member.churchMinistries || []).includes(ministry) ||
      (member.worshipServiceCount > 0 && ministry === '敬拜團');
    return matchesQuery &&
      matchesGroup &&
      matchesMinistry &&
      matchesAttendance(member.attendance && member.attendance.sunday, sunday) &&
      matchesAttendance(member.attendance && member.attendance.sundaySchool, sundaySchool) &&
      matchesAttendance(member.attendance && member.attendance.group, groupAttendance) &&
      matchesParticipation(member.participation, participation);
  });

  renderParticipationMap(rows);
  list.innerHTML = '';
  if (!rows.length) {
    list.innerHTML = '<p class="muted" style="padding:10px;">沒有符合條件的會友</p>';
    return;
  }

  rows.forEach(member => {
    const row = template.content.firstElementChild.cloneNode(true);
    row.dataset.uid = member.uid;
    if (member.uid === state.selectedUid) row.classList.add('active');
    row.querySelector('.member-name').textContent = member.name || '(未命名)';
    row.querySelector('.member-meta').textContent =
      `${member.uid || '無 UID'} · ${(member.groups || []).map(g => g.name).join('、') || '未分組'} · 參與 ${getParticipationCount(member)}`;
    row.addEventListener('click', () => selectMember(member.uid));
    list.appendChild(row);
  });
}

function matchesAttendance(bucket, filter) {
  if (!filter) return true;
  const count = bucket && bucket.count ? bucket.count : 0;
  const total = bucket && bucket.total ? bucket.total : 0;
  const rate = total > 0 ? Math.round((count / total) * 100) : 0;
  if (filter === 'present') return count > 0;
  if (filter === 'absent') return count === 0;
  if (filter === 'high') return total > 0 && rate >= 70;
  if (filter === 'low') return total > 0 && rate < 50;
  return true;
}

function matchesParticipation(participation, filter) {
  if (!filter) return true;
  return participation && participation.level === filter;
}

function renderParticipationMap(rows) {
  const container = document.getElementById('participationMap');
  const toggle = document.getElementById('participationToggle');
  document.getElementById('visibleCount').textContent = `${rows.length} 人`;
  const sorted = getVisibleParticipationRows(rows, state.participationExpanded);
  const canExpand = rows.length > 24;
  toggle.hidden = !canExpand;
  toggle.setAttribute('aria-expanded', String(canExpand && state.participationExpanded));
  toggle.textContent = state.participationExpanded ? '收起名單' : `顯示完整 ${rows.length} 人`;
  container.classList.toggle('expanded', canExpand && state.participationExpanded);
  if (!sorted.length) {
    container.innerHTML = '<p class="muted">沒有符合條件的會友</p>';
    return;
  }
  container.innerHTML = sorted.map(member => {
    const count = getParticipationCount(member);
    const dotClasses = getDotClasses(member);
    return `
      <button class="participation-row" type="button" data-uid="${escapeHtml(member.uid || '')}" title="${escapeHtml(member.name || '')}">
        <span class="participation-name">${escapeHtml(member.name || '')}</span>
        <span class="dot-track" aria-hidden="true">${Array.from({ length: 12 }).map((_, index) => {
          const isFilled = index < dotClasses.length;
          const fillClass = isFilled ? ` filled ${dotClasses[index]}` : '';
          return `<span class="dot${fillClass}"></span>`;
        }).join('')}</span>
        <span class="participation-score">${count}</span>
      </button>
    `;
  }).join('');
  container.querySelectorAll('.participation-row').forEach(row => {
    row.addEventListener('click', () => selectMember(row.dataset.uid));
  });
}

function getVisibleParticipationRows(rows, expanded) {
  const sorted = (rows || []).slice()
    .sort((a, b) => getParticipationCount(b) - getParticipationCount(a) || String(a.name || '').localeCompare(String(b.name || ''), 'zh-Hant'));
  return expanded ? sorted : sorted.slice(0, 24);
}

function getParticipationCount(member) {
  return member && member.participation ? Number(member.participation.ministryCount || 0) : 0;
}

async function selectMember(uid) {
  if (!uid) return;
  state.selectedUid = uid;
  renderMemberList();
  showBusy('載入會友詳情...');
  try {
    const result = await window.churchAPI('getProfile', { uid });
    if (!result || !result.success) throw new Error(result && result.message ? result.message : '讀取詳情失敗');
    renderProfile(result.profile, result.unresolvedParticipants || []);
  } catch (error) {
    renderError(error);
  } finally {
    hideBusy();
  }
}

function renderProfile(profile, unresolved) {
  const detail = document.getElementById('profileDetail');
  detail.innerHTML = `
    <div class="profile-header">
      <div class="profile-title">
        <h2>${escapeHtml(profile.name || '')}</h2>
        <span class="uid-pill">${escapeHtml(profile.uid || '')}</span>
      </div>
      <div class="muted">${escapeHtml(profile.gender || '')}</div>
    </div>

    <div class="section-grid">
      <section class="section">
        <h3>小組/團契</h3>
        ${renderGroups(profile.groups || [])}
      </section>

      <section class="section">
        <h3>門訓狀態</h3>
        <div class="tag-list">
          <span class="tag gold">保留入口</span>
          <span class="tag">${escapeHtml(profile.discipleship && profile.discipleship.status || 'unknown')}</span>
        </div>
      </section>

      <section class="section wide">
        <h3>近一年出席參考</h3>
        ${renderAttendance(profile.attendance || {})}
      </section>

      <section class="section wide">
        <h3>事工參與量</h3>
        ${renderParticipation(profile)}
      </section>

      <section class="section wide">
        <h3>小組/團契近一年服事</h3>
        ${renderGroupMinistries(profile.groupMinistries || [])}
      </section>

      <section class="section wide">
        <h3>未配對資料</h3>
        ${renderUnresolved(unresolved)}
      </section>
    </div>
  `;
}

function renderAttendance(attendance) {
  return `
    <div class="attendance-grid">
      ${renderAttendanceCard('主日禮拜', attendance.sunday)}
      ${renderAttendanceCard('主日學', attendance.sundaySchool)}
      ${renderAttendanceCard('小組聚會', attendance.group)}
    </div>
  `;
}

function renderAttendanceCard(label, bucket) {
  const count = bucket && bucket.count ? bucket.count : 0;
  const total = bucket && bucket.total ? bucket.total : 0;
  const rate = bucket && typeof bucket.rate === 'number' ? bucket.rate : 0;
  const lastDate = bucket && bucket.lastDate ? bucket.lastDate : '無';
  return `
    <div class="attendance-card">
      <span>${escapeHtml(label)}</span>
      <strong>${count}/${total}</strong>
      <small>${rate}% · 最近 ${escapeHtml(lastDate)}</small>
    </div>
  `;
}

function renderParticipation(profile) {
  const participation = profile && profile.participation ? profile.participation : {};
  const count = Number(participation.ministryCount || 0);
  const serviceCount = Number(participation.serviceCount || 0);
  const level = participation.level || 'none';
  const dots = Math.min(12, Math.max(0, count));

  // 取得所有參與的事工清單與類型
  const ministries = [];
  if (profile) {
    if (profile.groupMinistries && profile.groupMinistries.length) {
      profile.groupMinistries.forEach(item => {
        ministries.push({ name: item.groupName || item.ministryName, type: 'group' });
      });
    }
    if (profile.churchMinistries && profile.churchMinistries.length) {
      profile.churchMinistries.forEach(item => {
        ministries.push({ name: item.ministryName, type: 'ministry' });
      });
    }
    if (profile.worship && profile.worship.positions && profile.worship.positions.length) {
      profile.worship.positions.forEach(pos => {
        ministries.push({ name: `敬拜團 (${pos})`, type: 'ministry' });
      });
    }
  }

  const ministryTags = ministries.length
    ? `<div class="tag-list" style="margin-top:10px;">
        <span class="muted" style="font-size:13px; display:inline-flex; align-items:center;">參與項目：</span>
        ${ministries.map(m => {
          const tagClass = m.type === 'group' ? 'accent' : 'gold';
          return `<span class="tag ${tagClass}">${escapeHtml(m.name)}</span>`;
        }).join('')}
       </div>`
    : '';

  const dotClasses = getDotClasses(profile);

  return `
    <div class="tag-list" style="margin-bottom:10px;">
      <span class="tag accent">參與 ${count}</span>
      <span class="tag">服事紀錄 ${serviceCount}</span>
      <span class="tag gold">${escapeHtml(level)}</span>
    </div>
    <div class="dot-track" aria-label="事工參與量" style="margin-bottom:10px;">${Array.from({ length: 12 }).map((_, index) => {
      const isFilled = index < dotClasses.length;
      const fillClass = isFilled ? ` filled ${dotClasses[index]}` : '';
      return `<span class="dot${fillClass}"></span>`;
    }).join('')}</div>
    ${ministryTags}
  `;
}

function renderGroups(groups) {
  if (!groups.length) return '<p class="muted">未分屬小組/團契</p>';
  return `<div class="tag-list">${groups.map(g =>
    `<span class="tag accent">${escapeHtml(g.name)} · ${escapeHtml(g.role || '小羊')}</span>`
  ).join('')}</div>`;
}

function renderGroupMinistries(items) {
  if (!items.length) return '<p class="muted">近一年沒有小組/團契服事紀錄</p>';
  return items.map(item => `
    <div class="ministry-block" style="margin-bottom: 12px;">
      <h4 style="margin-bottom: 8px;">${escapeHtml(item.groupName)}</h4>
      <div class="tag-list">
        ${item.role ? `<span class="tag">${escapeHtml(item.role)}</span>` : ''}
        ${(item.duties || []).map(d => `<span class="tag gold">${escapeHtml(d)}</span>`).join('')}
      </div>
    </div>
  `).join('');
}





function renderHistory(items) {
  if (!items.length) return '<p class="muted">沒有紀錄</p>';
  return `<div class="history-list">${items
    .sort((a, b) => String(b.date).localeCompare(String(a.date)))
    .slice(0, 30)
    .map(item => `
      <div class="history-item">
        <span>${escapeHtml(item.date || '')}</span>
        <div><strong>${escapeHtml(item.label || '')}</strong><br>${escapeHtml(item.note || '')}</div>
      </div>
    `).join('')}</div>`;
}

function renderUnresolved(items) {
  if (!items.length) return '<p class="muted">此會友沒有相關未配對資料</p>';
  return renderHistory(items.map(item => ({
    date: item.date || '',
    label: `${item.name} · ${item.reason}`,
    note: `${item.source || ''} ${item.field || ''}`
  })));
}

function renderError(error) {
  const detail = document.getElementById('profileDetail');
  detail.innerHTML = `<div class="empty-state"><h2>讀取失敗</h2><p>${escapeHtml(error.message || String(error))}</p></div>`;
}

function showBusy(text) {
  document.getElementById('loadingText').textContent = text || '載入中...';
  document.getElementById('loadingOverlay').classList.add('show');
}

function hideBusy() {
  document.getElementById('loadingOverlay').classList.remove('show');
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getDotClasses(member) {
  if (!member) return [];
  const classes = [];

  // 1. Group ministries (聚會型)
  const groupCount = member.groupMinistries ? member.groupMinistries.length : 0;
  for (let i = 0; i < groupCount; i++) {
    classes.push('group-dot');
  }

  // 2. Church ministries (事工型)
  const churchCount = member.churchMinistries ? member.churchMinistries.length : 0;
  for (let i = 0; i < churchCount; i++) {
    classes.push('ministry-dot');
  }

  // 3. Worship team (事工型)
  let worshipCount = 0;
  if (member.worshipPositions) {
    worshipCount = member.worshipPositions.length;
  } else if (member.worship && member.worship.positions) {
    worshipCount = member.worship.positions.length;
  }
  for (let i = 0; i < worshipCount; i++) {
    classes.push('ministry-dot');
  }

  return classes.slice(0, 12);
}
