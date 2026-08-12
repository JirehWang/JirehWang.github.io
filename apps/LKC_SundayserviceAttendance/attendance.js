  var localPendingActions = {}; 
  var attUserId = localStorage.getItem('att_uid') || ('User_' + Math.floor(Math.random() * 1000000));
  localStorage.setItem('att_uid', attUserId);
  
  var attSyncTimer = null;
  var attTempTimer = null;
  var attTempUnsubscribe = null;
  var attTempModulePromise = null;
  var ATTENDANCE_TEMP_FLUSH_INTERVAL_MS = 30000;
  var ATTENDANCE_UI_QUEUE_KEY = 'attendance_ui_pending_v1';
  var MAX_ATTENDANCE_UI_QUEUE = 100;
  var attTempFlushInFlight = false;
  var attTempFlushPromise = null;
  var realtimeAttendanceTempEntries = {};
  var realtimeAttendanceTempPreviousEntries = {};
  var realtimeAttendanceTempDurableUntil = {};
  var realtimeAttendanceTempExpiryTimers = {};
  var realtimeAttendanceTempReady = false;
  var remoteStatusSequence = 0;
  var attIsRendering = false; 
  var currentAttType = '';
  var currentFormalRevision = 0;
  var html5QrCode = null; 
  var lastClickTime = 0; 
  var DOUBLE_CLICK_DELAY = 350; 
  var lastActiveTime = Date.now(); 
  var isSleeping = false; 
  var attSearchCacheKey = '';
  var attSearchCacheActive = false;
  var attSearchMemoryAnchor = null;
  
  var globalGroupConfig = {};

  function getAttendanceTempModule() {
    if (!attTempModulePromise) {
      var moduleUrl = new URL('../../firebase/attendance-temp.js', document.baseURI).href;
      attTempModulePromise = import(moduleUrl);
    }
    return attTempModulePromise;
  }

  function loadAttendanceTempQueue() {
    try {
      var parsed = JSON.parse(localStorage.getItem(ATTENDANCE_UI_QUEUE_KEY) || '[]');
      return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      return [];
    }
  }

  function saveAttendanceTempQueue(queue) {
    localStorage.setItem(ATTENDANCE_UI_QUEUE_KEY, JSON.stringify(queue.slice(-MAX_ATTENDANCE_UI_QUEUE)));
  }

  function getAttendanceTempDateValue(dateValue) {
    var dateInput = document.getElementById('attendanceDateInput');
    var raw = String(dateValue || (dateInput && dateInput.value) || '').trim();
    if (!raw) raw = formatDateToDash(new Date());
    return raw.replace(/\//g, '-');
  }

  function getAttendanceTempScope(scope, dateValue) {
    var base = String(scope || '').trim();
    if (!base || base.indexOf('AGM:') === 0 || base.indexOf('|date:') !== -1) return base;
    return base + '|date:' + getAttendanceTempDateValue(dateValue);
  }

  function getCurrentAttendanceTempScope() {
    return getAttendanceTempScope(currentAttType);
  }

  function enqueueAttendanceTemp(uid, checked, source) {
    var normalizedUid = String(uid || '').trim().toUpperCase();
    source = source === 'qr' ? 'qr' : 'manual';
    if (!/^LK\d+$/i.test(normalizedUid)) throw new Error('點名 UID 無效');
    var scope = getCurrentAttendanceTempScope();
    if (!scope) throw new Error('尚未選擇點名場次');
    var queue = loadAttendanceTempQueue();
    var item = queue.find(function(entry) { return entry.scope === scope && entry.uid === normalizedUid; });
    var now = Date.now();
    if (!item) {
      if (queue.length >= MAX_ATTENDANCE_UI_QUEUE) throw new Error('點名重試佇列已滿');
      var createdAt = now;
      item = {
        id: createdAt + '-' + Math.random().toString(36).slice(2),
        scope: scope,
        uid: normalizedUid,
        operatorId: attUserId,
        requestId: scope + ':' + normalizedUid + ':' + attUserId + ':' + createdAt,
        updatedAt: createdAt
      };
      queue.push(item);
    }
    item.checked = checked === true;
    item.source = source;
    item.updatedAt = now;
    item.requestId = scope + ':' + normalizedUid + ':' + attUserId + ':' + now;
    saveAttendanceTempQueue(queue);
    return item;
  }

  function updateAttendanceTempQueueItem(id, requestId, updatedAt, updater) {
    var latest = loadAttendanceTempQueue();
    var index = latest.findIndex(function(entry) {
      return entry.id === id && entry.requestId === requestId && entry.updatedAt === updatedAt;
    });
    if (index === -1) return false;
    var next = updater(latest[index]);
    if (next === null) latest.splice(index, 1);
    else if (next) latest[index] = next;
    saveAttendanceTempQueue(latest);
    return true;
  }

  function flushAttendanceTempQueue() {
    if (attTempFlushInFlight) return attTempFlushPromise || Promise.resolve({ attempted: 0, acknowledged: 0 });
    var queue = loadAttendanceTempQueue();
    if (!queue.length) return Promise.resolve({ attempted: 0, acknowledged: 0 });
    attTempFlushInFlight = true;
    var acknowledged = 0;
    var successfulWrites = 0;
    attTempFlushPromise = getAttendanceTempModule().then(function(store) {
      var batch = queue.slice(0, 12);
      return Promise.all(batch.map(function(item) {
        item.attempts = Number(item.attempts || 0) + 1;
        var writeRequestId = item.requestId;
        var writeUpdatedAt = item.updatedAt;
        return store.writeAttendanceTemp({
          scope: item.scope,
          uid: item.uid,
          checked: item.checked,
          operatorId: item.operatorId,
          source: item.source || 'manual',
          requestId: item.requestId,
          updatedAt: writeUpdatedAt
        }).then(function(result) {
          successfulWrites++;
          var finalEntry = result && typeof result === 'object' ? result : null;
          var durableKey = item.scope + '|' + item.uid;
          if (finalEntry && finalEntry.committed === false) {
            delete localPendingActions[item.uid];
            if (finalEntry.uid && finalEntry.checked !== undefined) {
              var finalEntries = {};
              finalEntries[item.uid] = finalEntry;
              applyRealtimeAttendanceTemp(finalEntries, { force: true });
            } else {
              clearRealtimeAttendanceTempCard(item.uid);
            }
          } else if (finalEntry && finalEntry.checked === true) {
            realtimeAttendanceTempDurableUntil[durableKey] = Number(finalEntry.expiresAt || (writeUpdatedAt + 6 * 60 * 60 * 1000));
          } else {
            delete realtimeAttendanceTempDurableUntil[durableKey];
          }
          if (updateAttendanceTempQueueItem(item.id, writeRequestId, writeUpdatedAt, function() { return null; })) {
            acknowledged++;
          }
        }).catch(function(error) {
          updateAttendanceTempQueueItem(item.id, writeRequestId, writeUpdatedAt, function(current) {
            return Object.assign({}, current, {
              attempts: Number(current.attempts || 0) + 1,
              lastError: String(error && error.message || error)
            });
          });
          console.warn('Firebase 點名暫存 ACK 失敗，保留重試', error);
        });
      })).then(function() {
        if (successfulWrites && loadAttendanceTempQueue().length) {
          setTimeout(function() { flushAttendanceTempQueue(); }, 0);
        }
        return { attempted: batch.length, acknowledged: acknowledged };
      });
    }).finally(function() {
      attTempFlushInFlight = false;
      attTempFlushPromise = null;
    });
    return attTempFlushPromise;
  }

  function flushAttendanceTempToBackend(scope) {
    flushAttendanceTempQueue().then(function() {
      return flushAttendanceTempToBackendAsync(scope);
    }).catch(function(error) {
      console.warn('Firebase attendance temp batch retry scheduled', error);
    });
  }

  function flushAttendanceTempToBackendAsync(scope) {
    if (!scope || typeof google === 'undefined' || !google.script || !google.script.run) {
      return Promise.resolve({ ok: true, skipped: true });
    }
    return new Promise(function(resolve, reject) {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .flushAttendanceTemp(scope, attUserId);
    });
  }

  function applyPendingSourceClass(card, entry) {
    if (!card) return;
    card.classList.remove('pending-manual', 'pending-qr');
    if (entry && entry.checked === true) {
      card.classList.add(entry.source === 'qr' ? 'pending-qr' : 'pending-manual');
    }
  }

  function clearRealtimeAttendanceTempCard(uid) {
    var container = document.getElementById('attendanceListBody');
    if (!container) return;
    var checkbox = container.querySelector('input[data-uid="' + uid + '"]');
    if (!checkbox) return;
    var card = checkbox.parentElement;
    if (!card || card.classList.contains('submitted')) return;
    checkbox.checked = false;
    checkbox.disabled = false;
    card.className = 'att-item shadow-sm';
    card.style.opacity = '1';
    card.style.pointerEvents = 'auto';
    applyPendingSourceClass(card, { checked: false });
    card.onclick = function(e) {
      e.preventDefault();
      var cb = this.querySelector('input');
      if (cb) { cb.checked = !cb.checked; toggleCardStyle(cb); }
    };
  }

  function scheduleRealtimeAttendanceTempExpiry(uid, entry, scope) {
    var expiresAt = Number(entry && entry.expiresAt || 0);
    if (!expiresAt) return;
    if (realtimeAttendanceTempExpiryTimers[uid]) clearTimeout(realtimeAttendanceTempExpiryTimers[uid]);
    var delay = Math.max(0, expiresAt - Date.now() + 25);
    realtimeAttendanceTempExpiryTimers[uid] = setTimeout(function() {
      delete realtimeAttendanceTempExpiryTimers[uid];
      var current = realtimeAttendanceTempEntries[uid];
      if (!current || Number(current.expiresAt || 0) <= Date.now()) {
        delete realtimeAttendanceTempDurableUntil[(scope || getCurrentAttendanceTempScope()) + '|' + uid];
        clearRealtimeAttendanceTempCard(uid);
      }
    }, Math.min(delay, 2147483647));
  }

  function applyRealtimeAttendanceTemp(entries, options) {
    var container = document.getElementById('attendanceListBody');
    if (!container || !entries || typeof entries !== 'object') return;
    var now = Date.now();
    var force = options && options.force === true;
    Object.keys(entries).forEach(function(uid) {
      var entry = entries[uid] || {};
      var checkbox = container.querySelector('input[data-uid="' + uid + '"]');
      if (!checkbox) return;
      var card = checkbox.parentElement;
      if (!card || card.classList.contains('submitted')) return;
      var pending = localPendingActions[uid];
      if (!force && pending && Number(pending.updatedAt || pending.time || 0) > Number(entry.updatedAt || 0)) return;
      var checked = entry.checked === true;
      var ownerId = entry.ownerId || entry.operatorId;
      var lockActive = Number(entry.lockedUntil || 0) > now && Number(entry.expiresAt || 0) > now;
      var lockedByOther = checked && lockActive && ownerId && ownerId !== attUserId;
      checkbox.checked = checked;
      checkbox.disabled = Boolean(lockedByOther);
      card.className = lockedByOther ? 'att-item shadow-sm locked' : (checked ? 'att-item shadow-sm selected' : 'att-item shadow-sm');
      applyPendingSourceClass(card, entry);
      card.style.opacity = lockedByOther ? '0.5' : '1';
      card.style.pointerEvents = lockedByOther ? 'none' : 'auto';
      card.onclick = lockedByOther ? null : function(e) {
        e.preventDefault();
        var cb = this.querySelector('input');
        if (cb) { cb.checked = !cb.checked; toggleCardStyle(cb); }
      };
      if (pending && Number(entry.updatedAt || 0) >= Number(pending.updatedAt || pending.time || 0)) {
        delete localPendingActions[uid];
      }
      scheduleRealtimeAttendanceTempExpiry(uid, entry);
    });
  }

  function startAttendanceTempSubscription(scope) {
    if (attTempUnsubscribe) attTempUnsubscribe();
    attTempUnsubscribe = null;
    realtimeAttendanceTempEntries = {};
    realtimeAttendanceTempPreviousEntries = {};
    realtimeAttendanceTempDurableUntil = {};
    Object.keys(realtimeAttendanceTempExpiryTimers).forEach(function(uid) {
      clearTimeout(realtimeAttendanceTempExpiryTimers[uid]);
    });
    realtimeAttendanceTempExpiryTimers = {};
    realtimeAttendanceTempReady = false;
    if (!scope) return;
    getAttendanceTempModule().then(function(store) {
      if (scope !== getCurrentAttendanceTempScope() || !store.subscribeAttendanceTemp) return;
      attTempUnsubscribe = store.subscribeAttendanceTemp(scope, function(entries) {
        if (scope !== getCurrentAttendanceTempScope()) return;
        var previous = realtimeAttendanceTempEntries || {};
        var next = entries && typeof entries === 'object' ? entries : {};
        Object.keys(previous).forEach(function(uid) {
          if (Object.prototype.hasOwnProperty.call(next, uid)) return;
          var durableUntil = Number(realtimeAttendanceTempDurableUntil[scope + '|' + uid] || previous[uid].expiresAt || 0);
          if (durableUntil > Date.now()) {
            scheduleRealtimeAttendanceTempExpiry(uid, { expiresAt: durableUntil }, scope);
          } else {
            clearRealtimeAttendanceTempCard(uid);
          }
        });
        realtimeAttendanceTempPreviousEntries = previous;
        realtimeAttendanceTempEntries = entries && typeof entries === 'object' ? entries : {};
        next = realtimeAttendanceTempEntries;
        realtimeAttendanceTempReady = true;
        applyRealtimeAttendanceTemp(realtimeAttendanceTempEntries);
      }, function(error) {
        realtimeAttendanceTempReady = false;
        console.warn('Firebase 點名即時同步暫時失敗，保留 GAS fallback', error);
      });
      realtimeAttendanceTempReady = true;
    }).catch(function(error) {
      realtimeAttendanceTempReady = false;
      console.warn('Firebase 點名即時同步模組載入失敗，保留 GAS fallback', error);
    });
  }

  function startAttendanceTempSync() {
    if (attTempTimer) clearInterval(attTempTimer);
    startAttendanceTempSubscription(getCurrentAttendanceTempScope());
    flushAttendanceTempQueue().then(function() {
      return flushAttendanceTempToBackendAsync(getCurrentAttendanceTempScope());
    }).catch(function(error) { console.warn('點名暫存批次同步稍後重試', error); });
    attTempTimer = setInterval(function() {
      flushAttendanceTempQueue().then(function() {
        return flushAttendanceTempToBackendAsync(getCurrentAttendanceTempScope());
      }).catch(function(error) { console.warn('點名暫存批次同步稍後重試', error); });
    }, ATTENDANCE_TEMP_FLUSH_INTERVAL_MS);
  }

  function stopAttendanceTempSync() {
    if (attTempTimer) clearInterval(attTempTimer);
    if (attTempUnsubscribe) attTempUnsubscribe();
    attTempTimer = null;
    attTempUnsubscribe = null;
    realtimeAttendanceTempEntries = {};
    realtimeAttendanceTempPreviousEntries = {};
    realtimeAttendanceTempDurableUntil = {};
    Object.keys(realtimeAttendanceTempExpiryTimers).forEach(function(uid) {
      clearTimeout(realtimeAttendanceTempExpiryTimers[uid]);
    });
    realtimeAttendanceTempExpiryTimers = {};
    realtimeAttendanceTempReady = false;
  }

  function updateBadgeUI(sleepMode) {
    var badge = document.getElementById('presentCountBadge');
    var pCountEl = document.getElementById('presentCount');
    var count = pCountEl ? pCountEl.innerText : '0';
    if (!badge) return;
    if (sleepMode) {
      badge.className = "badge p-2.5 small shadow-sm d-flex align-items-center flex-grow-1 justify-content-center border-0 sleep-pulse";
      badge.style.backgroundColor = "#4f46e5";
      badge.innerHTML = "💤 休眠中(點擊醒來) | 已出席：<span id='presentCount' class='mx-1'>" + count + "</span> 人";
    } else {
      badge.className = "badge p-2.5 small shadow-sm d-flex align-items-center flex-grow-1 justify-content-center border-0";
      badge.style.backgroundColor = "var(--color-green)";
      badge.style.boxShadow = "none";
      badge.innerHTML = "✅ 已出席：<span id='presentCount' class='mx-1'>" + count + "</span> 人";
    }
  }

  function wakeUp() {
    var wasSleeping = isSleeping;
    lastActiveTime = Date.now();
    if (wasSleeping) { isSleeping = false; updateBadgeUI(false); fetchRemoteStatus(); }
  }
  
  document.addEventListener('touchstart', wakeUp, { passive: true });
  document.addEventListener('mousemove', wakeUp, { passive: true });
  document.addEventListener('click', wakeUp, { passive: true });
  document.addEventListener('keydown', wakeUp, { passive: true });

  // 格式化：YYYY-MM-DD -> YYYY/M/D
  function formatDateToSlash(dateVal) {
    if (!dateVal) return "";
    const parts = dateVal.split('-');
    if (parts.length !== 3) return dateVal;
    const year = parts[0];
    const month = parseInt(parts[1], 10);
    const day = parseInt(parts[2], 10);
    return `${year}/${month}/${day}`;
  }

  // 格式化：Date對象 -> YYYY-MM-DD
  function formatDateToDash(dateObj) {
    const yyyy = dateObj.getFullYear();
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const dd = String(dateObj.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  // 解鎖日期編輯確認
  window.unlockDateEdit = function() {
    if (typeof autoJumpConfig !== 'undefined' && autoJumpConfig.active) {
      alert("⚠️ 透過場次 QR Code 進入，不可修改日期！");
      return;
    }
    if (confirm("確定要修改點名日期嗎？\n(這將會載入您所選擇日期的點名狀態，以避免誤觸)")) {
      const dateInput = document.getElementById('attendanceDateInput');
      if (dateInput) {
        dateInput.disabled = false;
        dateInput.focus();
      }
    }
  }

  // 檢查是否由 QR Code 進入並自動隱藏修改按鈕
  function checkDateLockStatus() {
    if (typeof autoJumpConfig !== 'undefined' && autoJumpConfig.active) {
      const unlockBtn = document.getElementById('unlockDateBtn');
      if (unlockBtn) {
        unlockBtn.style.display = 'none';
      }
      const dateInput = document.getElementById('attendanceDateInput');
      if (dateInput) {
        dateInput.disabled = true;
      }
    }
  }

  var today = new Date();
  var dateInput = document.getElementById('attendanceDateInput');
  if (dateInput) {
    dateInput.value = formatDateToDash(today);
    dateInput.addEventListener('change', function() {
      this.disabled = true; // 修改後立即重新鎖定
      if (currentAttType) {
        switchType(currentAttType); // 重新載入對應日期的名單
      }
    });
  }

  function loadGroupConfig(targetCategory = null, targetGroup = null) {
    google.script.run.withSuccessHandler(config => {
      globalGroupConfig = config;
      renderCategorySelect(targetCategory);
      updateGroupSelect(targetGroup);
      startAutoSync();
      checkDateLockStatus();
    }).getGroupConfig();
  }

  function renderCategorySelect(targetCategory) {
    const catSelect = document.getElementById('categorySelect');
    const modalCatSelect = document.getElementById('newGroupCategory');
    if (!catSelect) return;
    catSelect.innerHTML = '';
    if (modalCatSelect) modalCatSelect.innerHTML = '';
    for (let cat in globalGroupConfig) {
      let opt = document.createElement('option');
      opt.value = cat;
      opt.text = (cat === '禮拜' ? '⛪ ' : (cat === '主日學' ? '📖 ' : (cat === '禱告會' ? '🛐' : '📂 '))) + cat;
      if (cat === targetCategory) opt.selected = true;
      catSelect.add(opt);
      if (modalCatSelect) {
        let modalOpt = document.createElement('option');
        modalOpt.value = cat; modalOpt.text = opt.text;
        modalCatSelect.add(modalOpt);
      }
    }
  }

  function updateGroupSelect(targetGroup = null) {
    const catSelect = document.getElementById('categorySelect');
    const grpSelect = document.getElementById('groupSelect');
    if (!catSelect || !grpSelect) return;
    const currentCat = catSelect.value;
    grpSelect.innerHTML = '';
    if (globalGroupConfig[currentCat]) {
      globalGroupConfig[currentCat].forEach(grp => {
        let cleanGrp = grp.replace('點名紀錄', '').trim(); 
        let opt = document.createElement('option');
        opt.value = cleanGrp;
        opt.text = cleanGrp;
        if (cleanGrp === targetGroup || grp === targetGroup) opt.selected = true;
        grpSelect.add(opt);
      });
    }
    handleGroupSelect();
  }

  function handleGroupSelect() {
    const grpSelect = document.getElementById('groupSelect');
    if (grpSelect && grpSelect.value) switchType(grpSelect.value);
  }

  function switchType(type) {
    if (currentAttType === type && attIsRendering) return; 
    currentAttType = type;
    attIsRendering = true; 
    google.script.run.withFailureHandler(function(e){console.log("裝置綁定失敗",e)}).updateDeviceMode(attUserId, type);
    var listBody = document.getElementById('attendanceListBody');
    if (listBody) listBody.innerHTML = '<div class="full-width-msg"><div class="spinner-border text-primary mb-3"></div><div class="h6">讀取 [' + type + '] 名單中...</div></div>';
    var requestedType = type;
    
    var dateInput = document.getElementById('attendanceDateInput');
    var selectedDateStr = dateInput ? formatDateToSlash(dateInput.value) : "";

    google.script.run
      .withSuccessHandler(function(result) {
        if (requestedType !== currentAttType) return;
        var list = result.activeList || result;
        var nfMale = result.nfMale || 0;
        var nfFemale = result.nfFemale || 0;
        currentFormalRevision = Number(result.formalRevision || 0);
        renderAttendanceList(list, nfMale, nfFemale, currentFormalRevision);
        startAttendanceTempSubscription(getAttendanceTempScope(requestedType));
        setTimeout(function() { attIsRendering = false; }, 500);
      })
      .withFailureHandler(function(err){
        if (requestedType !== currentAttType) return; 
        alert("讀取名單失敗：" + err.message);
        attIsRendering = false;
      })
      .getSmartAttendanceList(requestedType, attUserId, selectedDateStr); 
  }

  function openGroupAddModal() {
    const modal = document.getElementById('groupAddModal');
    if (modal) modal.style.display = 'block';
    document.getElementById('newGroupName').value = '';
    const currentCat = document.getElementById('categorySelect').value;
    const modalCatSelect = document.getElementById('newGroupCategory');
    if (currentCat && modalCatSelect) modalCatSelect.value = currentCat;
  }

  function closeGroupAddModal() {
    const modal = document.getElementById('groupAddModal');
    if (modal) modal.style.display = 'none';
  }

  function saveNewGroup() {
    const category = document.getElementById('newGroupCategory').value;
    const groupName = document.getElementById('newGroupName').value.trim();
    const btn = document.getElementById('saveNewGroupBtn');
    if (!groupName) return alert("請輸入群組名稱！");
    if (btn) { btn.disabled = true; btn.innerText = "⏳ 建立中..."; }
    google.script.run
      .withFailureHandler(err => { alert(err.message); if (btn) { btn.disabled = false; btn.innerText = "確認新增"; } })
      .withSuccessHandler(config => {
        alert(`✅ 成功在 [${category}] 建立「${groupName}」！`);
        if (btn) { btn.disabled = false; btn.innerText = "確認新增"; }
        closeGroupAddModal();
        globalGroupConfig = config;
        renderCategorySelect(category);
        updateGroupSelect(groupName);
      })
      .createAttendanceGroup(category, groupName);
  }

  function renderAttendanceList(data, nfMale, nfFemale, formalRevision) {
    if (formalRevision !== undefined) currentFormalRevision = Number(formalRevision || 0);
    if (nfMale === undefined) nfMale = 0;
    if (nfFemale === undefined) nfFemale = 0;
    var container = document.getElementById('attendanceListBody');
    if (!container) return;
    container.innerHTML = ''; 
    if (!data || data.length === 0) {
      container.innerHTML = '<div class="full-width-msg p-5 h6">⚠️ 查無名單</div>';
      var pCount = document.getElementById('presentCount');
      if (pCount) pCount.innerText = '0';
      return;
    }
    var submittedCount = data.filter(function(m) { return m.isSubmitted; }).length;
    var presentCountEl = document.getElementById('presentCount');
    if (presentCountEl) presentCountEl.innerText = submittedCount + Number(nfMale) + Number(nfFemale);
    var maleEl = document.getElementById('newFriendsMale');
    var femaleEl = document.getElementById('newFriendsFemale');
    if (maleEl) maleEl.value = nfMale;
    if (femaleEl) femaleEl.value = nfFemale;
    var now = Date.now();
    data.forEach(function(m) {
      var label = document.createElement('label');
      label.className = "att-item shadow-sm";
      var lockedId = m.pendingOwnerId || m.ownerId || m.operatorId || m.userId || m.operator || m.uid;
      var isChecked = m.isChecked;
      var pendingLockUntil = Number(m.pendingLockedUntil || m.lockedUntil || 0);
      var isLocked = (m.isChecked && lockedId && lockedId !== attUserId
        && (!pendingLockUntil || pendingLockUntil > now));
      var isSubmitted = m.isSubmitted;
      var memberKey = m.id || m.name;
      label.dataset.scrollKey = encodeURIComponent(String(memberKey || ''));
      if (localPendingActions[memberKey]) {
        isChecked = localPendingActions[memberKey].state;
        isLocked = false;
      }
      var checkState = isChecked ? "checked" : "";
      var isDisabled = (isSubmitted || isLocked) ? "disabled" : "";
      var statusColor = m.gender === '男' ? '#0284c7' : (m.gender === '女' ? '#f43f5e' : 'var(--text-secondary)');
      var lockIcon = ""; 
      if (isSubmitted) {
        label.classList.add('submitted');
        statusColor = '#065f46';
        label.onclick = (function(memUid, memName) {
          return function(e) {
            e.preventDefault();
            var n = new Date().getTime();
            if (n - lastClickTime < DOUBLE_CLICK_DELAY) { confirmRevoke(memUid, memName); lastClickTime = 0; } else { lastClickTime = n; }
          };
        })(m.id, m.name);
      } else if (isLocked) {
        label.classList.add('locked');
        label.style.opacity = "0.5";
        label.style.pointerEvents = "none";
        lockIcon = ' <span style="font-size:12px;">🔒</span>'; 
      } else if (isChecked) {
        label.classList.add('selected');
        statusColor = 'rgba(255, 255, 255, 0.85)';
      } 
      applyPendingSourceClass(label, {
        checked: isChecked && !isSubmitted,
        source: m.pendingSource || m.source || 'manual'
      });
      if (!isSubmitted && !isLocked) {
        label.onclick = function(e) {
          e.preventDefault();
          var cb = this.querySelector('input');
          if (cb) { cb.checked = !cb.checked; toggleCardStyle(cb); }
        };
      }
      var uidString = m.id ? m.id : '';
      var genderString = m.gender ? m.gender : '未知';
      label.innerHTML = 
        '<input type="checkbox" value="' + m.name + '" data-uid="' + uidString + '" ' + checkState + ' ' + isDisabled + '>' +
        '<div class="att-name">' + m.name + lockIcon + '</div>' +
        '<div class="att-info"><b class="gender-text" style="color: ' + statusColor + ';">' + genderString + '</b></div>';
      container.appendChild(label);
    });
  }

  function toggleCardStyle(checkbox) {
    var isChecked = checkbox.checked;
    var uid = checkbox.dataset.uid || checkbox.value;
    if (isChecked) checkbox.parentElement.classList.add('selected');
    else checkbox.parentElement.classList.remove('selected');
    localPendingActions[uid] = { time: Date.now(), updatedAt: Date.now(), state: isChecked, source: 'manual' };
    if (!/^LK\d+$/i.test(String(uid || '').trim())) {
      google.script.run.withFailureHandler(function(err) {
        checkbox.checked = !isChecked;
        checkbox.parentElement.classList.toggle('selected');
        delete localPendingActions[uid];
      }).syncClickToServer(uid, isChecked, currentAttType, attUserId);
      return;
    }
    try {
      enqueueAttendanceTemp(uid, isChecked, 'manual');
      flushAttendanceTempQueue().catch(function(error) { console.warn('點名暫存稍後重試', error); });
    } catch (error) {
      checkbox.checked = !isChecked;
      checkbox.parentElement.classList.toggle('selected');
      delete localPendingActions[uid];
      console.warn('點名暫存建立失敗', error);
    }
  }

  function openAttendanceAddModal() {
    var modal = document.getElementById('attendanceAddModal');
    if (modal) modal.style.display = 'block';

    var nameInput = document.getElementById('editName_Att');
    var genderSelect = document.getElementById('editGender_Att');
    var noteInput = document.getElementById('editNote_Att');
    var excludedInput = document.getElementById('editIsExcluded_Att');
    if (nameInput) nameInput.value = '';
    if (genderSelect) genderSelect.selectedIndex = 0;
    if (noteInput) noteInput.value = '';
    if (excludedInput) excludedInput.checked = false;
  }

  function closeAttendanceAddModal() {
    var modal = document.getElementById('attendanceAddModal');
    if (modal) modal.style.display = 'none';
  }

  function saveNewMemberFromAttendance() {
    var btn = document.getElementById('saveNewMemberBtn_Att');
    var newData = {
      name: document.getElementById('editName_Att').value.trim(),
      gender: document.getElementById('editGender_Att').value,
      note: document.getElementById('editNote_Att').value.trim(),
      isExcluded: document.getElementById('editIsExcluded_Att').checked
    };
    if (!newData.name) return alert("請輸入姓名！");
    if (btn) { btn.disabled = true; btn.innerText = "儲存中..."; }
    google.script.run.withSuccessHandler(function(msg) {
          if (btn) { btn.disabled = false; btn.innerText = "確認儲存"; }
          if (msg.includes("成功")) { closeAttendanceAddModal(); setTimeout(function(){ silentRefreshList(); }, 600); } 
          else { alert(msg); }
      }).addMember(newData); 
  }

  function silentRefreshList() {
    if (attIsRendering) return;
    attIsRendering = true;
    var dateInput = document.getElementById('attendanceDateInput');
    var selectedDateStr = dateInput ? formatDateToSlash(dateInput.value) : "";
    google.script.run.withSuccessHandler(function(result) {
        var list = Array.isArray(result) ? result : (result.activeList || []);
        var nfMale = result.nfMale || 0;
        var nfFemale = result.nfFemale || 0;
        currentFormalRevision = Number(result.formalRevision || 0);
        renderAttendanceList(list, nfMale, nfFemale, currentFormalRevision);
        attIsRendering = false;
    }).getSmartAttendanceList(currentAttType, attUserId, selectedDateStr);
  }

  function confirmRevoke(uid, displayName) {
    if (navigator.vibrate) navigator.vibrate(50);
    if (confirm("確定要撤銷 [" + (displayName || uid) + "] 的送出紀錄嗎？")) { executeRevoke(uid, displayName); }
  }

  function writeAttendanceTempTombstone(uid, scope) {
    if (!/^LK\d+$/i.test(String(uid || '').trim())) return Promise.resolve({ committed: true, skipped: true });
    return getAttendanceTempModule().then(function(store) {
      return store.writeAttendanceTemp({
        scope: scope,
        uid: uid,
        checked: false,
        operatorId: attUserId,
        source: 'manual',
        requestId: 'revoke:' + scope + ':' + uid + ':' + Date.now(),
        updatedAt: Date.now()
      });
    }).then(function(result) {
      if (result && result.committed === false) throw new Error('此筆預點名正由其他裝置更新，請稍後再試');
      return result;
    });
  }

function executeRevoke(uid, displayName) {
    var btn = document.getElementById('submitBtn');
    var originalText = "確認送出";
    if (btn) { btn.disabled = true; btn.innerHTML = '正在撤銷...'; }
    var dateInput = document.getElementById('attendanceDateInput');
    var selectedDateStr = dateInput ? formatDateToSlash(dateInput.value) : "";
    writeAttendanceTempTombstone(uid, getCurrentAttendanceTempScope()).then(function() {
      return new Promise(function(resolve, reject) {
        google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)
          .revokeAttendance(uid, currentAttType, attUserId, selectedDateStr, currentFormalRevision);
      });
    }).then(function(msg) {
        if (msg === 'STALE_REVISION') {
          alert('已被其他裝置更新，畫面將重新整理');
          switchType(currentAttType);
          return;
        }
        if (msg === "OK") {
            var container = document.getElementById('attendanceListBody');
            var checkbox = container.querySelector('input[data-uid="' + uid + '"]');
            if (checkbox) {
                var card = checkbox.parentElement;
                card.className = "att-item shadow-sm";
                card.style.opacity = "1";
                card.style.pointerEvents = "auto";
                checkbox.checked = false;
                checkbox.disabled = false;
                var nameDiv = card.querySelector('.att-name');
                if (nameDiv) nameDiv.innerHTML = displayName || uid;
                card.onclick = function(e) {
                    e.preventDefault();
                    var cb = this.querySelector('input');
                    cb.checked = !cb.checked;
                    toggleCardStyle(cb);
                };
            }
            localPendingActions[uid] = { time: Date.now(), updatedAt: Date.now(), state: false };
            alert("✅ 撤銷成功！");
            var submittedCards = container.querySelectorAll('.att-item.submitted').length;
            var maleEl = document.getElementById('newFriendsMale');
            var femaleEl = document.getElementById('newFriendsFemale');
            var nfMale = maleEl ? parseInt(maleEl.value) || 0 : 0;
            var nfFemale = femaleEl ? parseInt(femaleEl.value) || 0 : 0;
            var presentEl = document.getElementById('presentCount');
            if (presentEl) presentEl.innerText = submittedCards + nfMale + nfFemale;
        } else {
            alert(msg);
        }
        if (btn) { btn.disabled = false; btn.innerHTML = originalText; }
    }).catch(function(error) {
      alert(error && error.message ? error.message : error);
      if (btn) { btn.disabled = false; btn.innerHTML = originalText; }
    });
}

  function fetchRemoteStatus() {
    var searchInput = document.getElementById('attSearchInput');
    if (attIsRendering || (searchInput && searchInput.value)) return;
    var dateInput = document.getElementById('attendanceDateInput');
    var selectedDateStr = dateInput ? formatDateToSlash(dateInput.value) : "";
    var requestSequence = ++remoteStatusSequence;
    var requestScope = currentAttType;
    google.script.run.withSuccessHandler(function(data) {
        if (!data || attIsRendering || requestSequence !== remoteStatusSequence || requestScope !== currentAttType) return;
        var activeList = Array.isArray(data) ? data : (data.activeList || []);
        var nfMale = Array.isArray(data) ? 0 : (data.nfMale || 0);
        var nfFemale = Array.isArray(data) ? 0 : (data.nfFemale || 0);
        var now = Date.now();
        var container = document.getElementById('attendanceListBody');
        if (!container) return;
        activeList.forEach(function(m) {
           var checkbox = container.querySelector('input[data-uid="' + m.id + '"]');
           if (!checkbox) return;
           var card = checkbox.parentElement;
           var memKey = m.id || m.name;
           if (localPendingActions[memKey] && Number(localPendingActions[memKey].updatedAt || localPendingActions[memKey].time || 0) >= Number(m.pendingUpdatedAt || 0)) return;
           var lockedId = m.pendingOwnerId || m.ownerId || m.operatorId || m.userId || m.operator || m.uid;
           var pendingLockUntil = Number(m.pendingLockedUntil || m.lockedUntil || 0);
           var pendingExpiresAt = Number(m.pendingExpiresAt || m.expiresAt || 0);
           var pendingSource = m.pendingSource || m.source || 'manual';
           applyPendingSourceClass(card, {
             checked: !m.isSubmitted && m.isChecked,
             source: pendingSource
           });
           if (m.isSubmitted) {
             if (!card.classList.contains('submitted')) {
                card.className = "att-item shadow-sm submitted";
                checkbox.checked = true; checkbox.disabled = true;
                card.style.opacity = "1"; card.style.pointerEvents = "auto";
                var nameDiv = card.querySelector('.att-name');
                if (nameDiv) nameDiv.innerHTML = m.name;
                card.onclick = (function(memUid, memName) {
                  return function(e) { e.preventDefault();
                    var n = new Date().getTime();
                    if (n - lastClickTime < DOUBLE_CLICK_DELAY) { confirmRevoke(memUid, memName); lastClickTime = 0; } else { lastClickTime = n; }
                  };
                })(m.id, m.name);
             }
           } else if (realtimeAttendanceTempReady || Object.prototype.hasOwnProperty.call(realtimeAttendanceTempEntries, memKey)) {
              // Firebase owns all non-submitted pending state. A slower GAS response
              // must never roll a realtime checked value back to its stale Sheet value.
              return;
           } else if (m.isChecked) {
             checkbox.checked = true;
              if (lockedId && lockedId !== attUserId
                  && (!pendingLockUntil || pendingLockUntil > now)
                  && (!pendingExpiresAt || pendingExpiresAt > now)) {
               card.className = "att-item shadow-sm locked"; 
               checkbox.disabled = true; card.onclick = null;
               card.style.opacity = "0.5"; card.style.pointerEvents = "none";
               var nameDiv = card.querySelector('.att-name');
               if (nameDiv && nameDiv.innerHTML.indexOf('🔒') === -1) nameDiv.innerHTML += ' <span style="font-size:12px;">🔒</span>';
             } else {
               card.className = "att-item shadow-sm selected"; 
               checkbox.disabled = false;
               card.style.opacity = "1"; card.style.pointerEvents = "auto";
               var nameDiv = card.querySelector('.att-name');
               if (nameDiv) nameDiv.innerHTML = m.name;
               card.onclick = function(e) { e.preventDefault(); var cb=this.querySelector('input'); cb.checked=!cb.checked; toggleCardStyle(cb); };
             }
           } else {
             card.className = "att-item shadow-sm"; 
             checkbox.checked = false; checkbox.disabled = false;
             card.style.opacity = "1"; card.style.pointerEvents = "auto";
             var nameDiv = card.querySelector('.att-name');
             if (nameDiv) nameDiv.innerHTML = m.name;
             card.onclick = function(e) { e.preventDefault(); var cb = this.querySelector('input'); cb.checked = !cb.checked; toggleCardStyle(cb); };
           }
        });
        var submittedCards = container.querySelectorAll('.att-item.submitted').length;
        var presentEl = document.getElementById('presentCount');
        if (presentEl) presentEl.innerText = submittedCards + Number(nfMale) + Number(nfFemale);
    }).getQuickSyncData(currentAttType, attUserId, selectedDateStr);
  }
  
  function startAutoSync() { 
    stopAutoSync(); 
    startAttendanceTempSync();
    attSyncTimer = setInterval(function() {
      if (Date.now() - lastActiveTime > 20000) {
        if (!isSleeping) { isSleeping = true; updateBadgeUI(true); }
        return; 
      }
      fetchRemoteStatus();
    }, 10000); 
  }
  
  function stopAutoSync() { if (attSyncTimer) clearInterval(attSyncTimer); stopAttendanceTempSync(); }

  async function submitAttendance() {
    var checked = document.querySelectorAll('.att-item.selected input:checked:not([disabled])');
    var memberCount = checked.length;
    var maleEl = document.getElementById('newFriendsMale');
    var femaleEl = document.getElementById('newFriendsFemale');
    var maleCount = maleEl ? parseInt(maleEl.value) || 0 : 0;
    var femaleCount = femaleEl ? parseInt(femaleEl.value) || 0 : 0;
    var totalNf = maleCount + femaleCount;
    var grandTotal = memberCount + totalNf;
    if (grandTotal === 0) {
        if (!confirm("⚠️ 目前沒有勾選任何名單，新朋友也是 0 人。\n確定要執行送出嗎？(這將會把新朋友數量歸零)")) return;
    } else {
        var confirmMsg = "📊 本次點名總計：" + grandTotal + " 人\n" +
                         "------------------------\n" +
                         "✅ 正式會友：" + memberCount + " 人\n" +
                         "✨ 新朋友：" + totalNf + " 人 (男:" + maleCount + " / 女:" + femaleCount + ")\n" +
                         "------------------------\n" +
                         "確定要正式送出紀錄嗎？";
        if (!confirm(confirmMsg)) return;
    }
    // 送出 UID 列表（後端會用主日 cache 反查姓名/性別，前端不再傳遞 (男)/(女) 後綴）
    var presentList = Array.from(checked).map(function(cb) {
       return cb.dataset.uid || cb.value;
    });
    var btn = document.getElementById('submitBtn');
    var originalText = "確認送出";
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> 處理中...'; }
    try {
      await flushAttendanceTempQueue();
      await flushAttendanceTempToBackendAsync(getCurrentAttendanceTempScope());
    } catch (error) {
      alert('點名暫存尚未取得 Firebase ACK，請確認網路後重試。');
      if (btn) { btn.disabled = false; btn.innerHTML = originalText; }
      return;
    }
    attIsRendering = true; 
    var dateInput = document.getElementById('attendanceDateInput');
    var dateText = dateInput ? formatDateToSlash(dateInput.value) : '';
      google.script.run.withSuccessHandler(function(msg) {
        if (msg === 'STALE_REVISION') {
          alert('已被其他裝置更新，畫面將重新整理');
        } else {
          alert(msg);
        }
        if (btn) { btn.disabled = false; btn.innerHTML = originalText; }
        attIsRendering = false; 
        switchType(currentAttType); 
      }).withFailureHandler(function(err) {
        alert("送出失敗：" + err.message);
        if (btn) { btn.disabled = false; btn.innerHTML = originalText; }
        attIsRendering = false;
      }).saveAttendance(dateText, presentList, currentAttType, maleCount, femaleCount, currentFormalRevision);
  }

  function filterAttList() {
    var searchInput = document.getElementById('attSearchInput');
    if (!searchInput) return;
    var kw = searchInput.value.trim().toLowerCase();
    var items = Array.prototype.slice.call(document.querySelectorAll('label.att-item'));
    var scrollArea = document.querySelector('.attendance-scroll-area');
    var anchor = null;
    var cacheKey = window.AttendanceSearchScroll
      ? window.AttendanceSearchScroll.getKey(currentAttType, document.getElementById('attendanceDateInput')?.value)
      : '';

    if (kw !== "" && scrollArea && window.ListScrollAnchor && window.AttendanceSearchScroll) {
      if (!attSearchCacheActive || attSearchCacheKey !== cacheKey) {
        var preSearchAnchor = window.ListScrollAnchor.capture(scrollArea, 'label.att-item');
        var storage = null;
        try { storage = window.localStorage; } catch (error) { storage = null; }
        var stored = window.AttendanceSearchScroll.save(storage, cacheKey, preSearchAnchor);
        attSearchMemoryAnchor = stored ? null : preSearchAnchor;
        attSearchCacheKey = cacheKey;
        attSearchCacheActive = true;
      }
    } else if (kw === "" && attSearchCacheActive && attSearchCacheKey === cacheKey && window.AttendanceSearchScroll) {
      var restoreStorage = null;
      try { restoreStorage = window.localStorage; } catch (error) { restoreStorage = null; }
      anchor = window.AttendanceSearchScroll.consume(restoreStorage, cacheKey) || attSearchMemoryAnchor;
      attSearchMemoryAnchor = null;
      attSearchCacheKey = '';
      attSearchCacheActive = false;
    }

    items.forEach(function(item) {
      var nameEl = item.querySelector('.att-name');
      if (nameEl) {
          var name = nameEl.innerText.trim().toLowerCase();
          item.style.display = (kw === "" || name.includes(kw)) ? 'flex' : 'none';
      }
    });

    if (anchor && scrollArea && window.ListScrollAnchor) {
      var restoreAnchor = function() {
        window.ListScrollAnchor.restore(scrollArea, 'label.att-item', anchor);
      };
      if (typeof requestAnimationFrame === 'function') requestAnimationFrame(restoreAnchor);
      else setTimeout(restoreAnchor, 0);
    }
  }

  function toggleScanner() {
    if (!attUserId) attUserId = localStorage.getItem('att_uid');
    localStorage.setItem('attendance_scope', currentAttType || '');
    var scannerUrl = "https://jirehwang.github.io/LKC1958_June_1.github.io/apps/qrcodescanner.github.io/?mode=" + encodeURIComponent(currentAttType || '') + "&date=" + encodeURIComponent(getAttendanceTempDateValue()) + "&context=attendance";
    window.open(scannerUrl, '_blank');
    startAutoSync();
  }

  function startScanning() {
    html5QrCode = new Html5Qrcode("reader");
    Html5Qrcode.getCameras().then(function(devices) {
      if (devices && devices.length > 0) {
        var cameraId = devices.length > 1 ? devices[devices.length - 1].id : devices[0].id;
        html5QrCode.start(
          cameraId, { fps: 10, qrbox: { width: 250, height: 250 } }, 
          function(decodedText) { handleQrCodeResult(decodedText.trim()); }
        ).catch(function(err) { alert("相機啟動失敗：" + err); toggleScanner(); });
      } else { alert("找不到攝影機設備！"); toggleScanner(); }
    }).catch(function(err) { alert("無法取得相機權限：" + err); toggleScanner(); });
  }

  function stopScanning() { if (html5QrCode && html5QrCode.isScanning) { html5QrCode.stop().then(function(){ html5QrCode.clear(); }); } }

  function handleQrCodeResult(scannedText) {
    wakeUp(); 
    var found = false;
    var checkboxes = document.querySelectorAll('#attendanceListBody input[type="checkbox"]');
    for (var i = 0; i < checkboxes.length; i++) {
      var cb = checkboxes[i];
      if ((cb.dataset.uid === scannedText || cb.value === scannedText) && !cb.disabled) {
        if (!cb.checked) { cb.checked = true; toggleCardStyle(cb); cb.parentElement.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
        if (navigator.vibrate) navigator.vibrate(100); 
        found = true; break;
      }
    }
    if (!found) console.log("⚠️ 掃描成功但名單中找不到此人：" + scannedText);
  }

  function downloadVenueJumpCard(cat, grp) {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = 400; canvas.height = 640;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, 400, 640);
    ctx.strokeStyle = '#D4AF37';
    ctx.lineWidth = 10;
    ctx.strokeRect(15, 15, 370, 610);
    ctx.textAlign = 'center';
    ctx.fillStyle = '#2c3e50';
    ctx.font = 'bold 30px "Microsoft JhengHei"';
    ctx.fillText('林口長老教會', 200, 80);
    ctx.font = 'bold 36px "Microsoft JhengHei"';
    ctx.fillText(grp + ' 點名處', 200, 130);
  
    // ✅ 加上時間戳記，強制瀏覽器視為全新請求
    const timestamp = Date.now();
    const finalUrl = "https://jirehwang.github.io/LKC1958_June_1.github.io/apps/LKC_SundayserviceAttendance/?cat=" + encodeURIComponent(cat) + "&grp=" + encodeURIComponent(grp) + "&v=" + timestamp;
    
    const qr = new QRious({ value: finalUrl, size: 300, level: 'H' });
    ctx.drawImage(qr.canvas, 50, 180, 300, 300);
    ctx.fillStyle = '#7f8c8d';
    ctx.font = '18px "Microsoft JhengHei"';
    ctx.fillText('使用手機相機掃描', 200, 520);
    ctx.fillText('即可直接開啟此場次名單', 200, 550);
    ctx.fillStyle = '#2c3e50';
    ctx.font = 'bold 20px "Consolas"';
    ctx.fillText('AUTO JUMP SYSTEM', 200, 590);
    const link = document.createElement('a');
    link.download = "場次自動跳轉卡_" + grp + ".png";
    link.href = canvas.toDataURL('image/png');
    link.click();
  }

  function triggerVenueDownload() {
    const cat = document.getElementById('categorySelect').value;
    const grp = document.getElementById('groupSelect').value;
    if (!cat || !grp || grp === "載入中...") { alert("⚠️ 請先選擇點名類別與群組！"); return; }
    downloadVenueJumpCard(cat, grp);
  }

  // 頁面載入時初始化分支判斷
  const params = new URLSearchParams(window.location.search);
  const initCat = params.get('cat');
  const initGrp = params.get('grp');

  if (initCat && initGrp) {
    // 啟動鎖定點名模式
    currentAttType = initGrp;
    attIsRendering = true;

    // 1. 顯示鎖定 Banner，隱藏選單列、管理按鈕與回主選單按鈕
    const lockBanner = document.getElementById('lockModeBanner');
    if (lockBanner) lockBanner.style.display = 'block';
    const lockText = document.getElementById('lockModeText');
    if (lockText) lockText.innerText = `${initCat} - ${initGrp}`;

    const groupSelectRow = document.getElementById('groupSelectRow');
    if (groupSelectRow) groupSelectRow.style.display = 'none';

    const btnGroupAdd = document.getElementById('btnGroupAdd');
    if (btnGroupAdd) btnGroupAdd.style.display = 'none';

    // 隱藏下載場次 QR 按鈕
    const btnDownloadVenue = document.querySelector('button[onclick="triggerVenueDownload()"]');
    if (btnDownloadVenue) btnDownloadVenue.style.display = 'none';

    // 隱藏 index.html 中的「← 回主選單」按鈕
    const backBtn = document.querySelector('#content-area > button[onclick="showHome()"]');
    if (backBtn) backBtn.style.display = 'none';

    // 2. 顯示讀取中，立即抓取名單與裝置綁定 (只有 1 個 API 請求)
    google.script.run.withFailureHandler(e => console.log("裝置綁定失敗", e)).updateDeviceMode(attUserId, initGrp);
    
    var listBody = document.getElementById('attendanceListBody');
    if (listBody) {
      listBody.innerHTML = '<div class="full-width-msg"><div class="spinner-border text-primary mb-3"></div><div class="h6">讀取 [' + initGrp + '] 名單中...</div></div>';
    }
    
    google.script.run
      .withSuccessHandler(result => {
        var list = result.activeList || result;
        var nfMale = result.nfMale || 0;
        var nfFemale = result.nfFemale || 0;
        currentFormalRevision = Number(result.formalRevision || 0);
        renderAttendanceList(list, nfMale, nfFemale, currentFormalRevision);
        startAttendanceTempSubscription(getAttendanceTempScope(initGrp));
        setTimeout(() => { attIsRendering = false; }, 500);
      })
      .withFailureHandler(err => {
        alert("讀取名單失敗：" + err.message);
        attIsRendering = false;
      })
      .getSmartAttendanceList(initGrp, attUserId);
      
    startAutoSync();
  } else {
    // 正常模式：載入選單與設定
    loadGroupConfig();
  }
