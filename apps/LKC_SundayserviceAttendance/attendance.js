  var localPendingActions = {}; 
  var attUserId = localStorage.getItem('att_uid') || ('User_' + Math.floor(Math.random() * 1000000));
  localStorage.setItem('att_uid', attUserId);
  
  var attSyncTimer = null;
  var attIsRendering = false; 
  var currentAttType = '';
  var html5QrCode = null; 
  var lastClickTime = 0; 
  var DOUBLE_CLICK_DELAY = 350; 
  var lastActiveTime = Date.now(); 
  var isSleeping = false; 
  
  var globalGroupConfig = {};

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
        renderAttendanceList(list, nfMale, nfFemale);
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

  function renderAttendanceList(data, nfMale, nfFemale) {
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
      var lockedId = m.operatorId || m.userId || m.operator || m.uid;
      var isChecked = m.isChecked;
      var isLocked = (m.isChecked && lockedId && lockedId !== attUserId);
      var isSubmitted = m.isSubmitted;
      var memberKey = m.id || m.name;
      label.dataset.scrollKey = encodeURIComponent(String(memberKey || ''));
      if (localPendingActions[memberKey] && (now - localPendingActions[memberKey].time < 5000)) {
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

    updateAgmQuorumBanner();
  }

  // 🏛️ 和會/會員大會成會即時計算器 (針對 204 名應到會員)
  function updateAgmQuorumBanner() {
    var banner = document.getElementById('agmQuorumBanner');
    if (!banner) return;

    var officialList = window.INITIAL_OFFICIAL_MEMBERS || [];
    var activeCommunicants = officialList.filter(function(m) {
      return (m.category_code || m.categoryCode) === 'CAT_1';
    });

    var totalActiveCount = activeCommunicants.length || 204;
    var threshold = Math.ceil(totalActiveCount * 0.5);

    var activeNamesMap = {};
    activeCommunicants.forEach(function(m) { activeNamesMap[m.name] = true; });

    var presentCount = 0;
    var checkboxes = document.querySelectorAll('#attendanceListBody input[type="checkbox"]');
    checkboxes.forEach(function(cb) {
      if ((cb.checked || cb.parentElement.classList.contains('submitted')) && activeNamesMap[cb.value]) {
        presentCount++;
      }
    });

    var percent = Math.min(100, Math.round((presentCount / totalActiveCount) * 100));
    var progressBar = document.getElementById('agmProgressBar');
    var presentEl = document.getElementById('agmPresentCount');
    var statusBadge = document.getElementById('agmQuorumStatusBadge');

    if (progressBar) progressBar.style.width = percent + '%';
    if (presentEl) presentEl.innerText = presentCount;

    if (statusBadge) {
      if (presentCount >= threshold) {
        statusBadge.className = "badge bg-success text-white fw-bold px-2 py-1";
        statusBadge.innerText = "✅ 已達 50% 成會門檻 (" + percent + "%)";
      } else {
        var needed = threshold - presentCount;
        statusBadge.className = "badge bg-warning text-dark fw-bold px-2 py-1";
        statusBadge.innerText = "⚠️ 尚差 " + needed + " 人成會 (" + percent + "%)";
      }
    }
  }

  function toggleCardStyle(checkbox) {
    var isChecked = checkbox.checked;
    var uid = checkbox.dataset.uid || checkbox.value;
    if (isChecked) checkbox.parentElement.classList.add('selected');
    else checkbox.parentElement.classList.remove('selected');
    updateAgmQuorumBanner();
    localPendingActions[uid] = { time: Date.now(), state: isChecked };
    google.script.run.withFailureHandler(function(err) {
        checkbox.checked = !isChecked;
        checkbox.parentElement.classList.toggle('selected');
        delete localPendingActions[uid];
        updateAgmQuorumBanner();
    }).syncClickToServer(uid, isChecked, currentAttType, attUserId);
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
        renderAttendanceList(list, nfMale, nfFemale);
        attIsRendering = false;
    }).getSmartAttendanceList(currentAttType, attUserId, selectedDateStr);
  }

  function confirmRevoke(uid, displayName) {
    if (navigator.vibrate) navigator.vibrate(50);
    if (confirm("確定要撤銷 [" + (displayName || uid) + "] 的送出紀錄嗎？")) { executeRevoke(uid, displayName); }
  }

function executeRevoke(uid, displayName) {
    var btn = document.getElementById('submitBtn');
    var originalText = "確認送出";
    if (btn) { btn.disabled = true; btn.innerHTML = '正在撤銷...'; }
    var dateInput = document.getElementById('attendanceDateInput');
    var selectedDateStr = dateInput ? formatDateToSlash(dateInput.value) : "";
    google.script.run.withSuccessHandler(function(msg) {
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
            localPendingActions[uid] = { time: Date.now(), state: false };
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
    }).revokeAttendance(uid, currentAttType, attUserId, selectedDateStr);
}

  function fetchRemoteStatus() {
    var searchInput = document.getElementById('attSearchInput');
    if (attIsRendering || (searchInput && searchInput.value)) return;
    var dateInput = document.getElementById('attendanceDateInput');
    var selectedDateStr = dateInput ? formatDateToSlash(dateInput.value) : "";
    google.script.run.withSuccessHandler(function(data) {
        if (!data || attIsRendering) return;
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
           if (localPendingActions[memKey] && (now - localPendingActions[memKey].time < 5000)) return;
           var lockedId = m.operatorId || m.userId || m.operator || m.uid;
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
           } else if (m.isChecked) {
             checkbox.checked = true;
             if (lockedId && lockedId !== attUserId) {
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
    attSyncTimer = setInterval(function() {
      if (Date.now() - lastActiveTime > 20000) {
        if (!isSleeping) { isSleeping = true; updateBadgeUI(true); }
        return; 
      }
      fetchRemoteStatus();
    }, 10000); 
  }
  
  function stopAutoSync() { if (attSyncTimer) clearInterval(attSyncTimer); }

  function submitAttendance() {
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
    attIsRendering = true; 
    var dateInput = document.getElementById('attendanceDateInput');
    var dateText = dateInput ? formatDateToSlash(dateInput.value) : '';
    google.script.run.withSuccessHandler(function(msg) {
        alert(msg); 
        if (btn) { btn.disabled = false; btn.innerHTML = originalText; }
        attIsRendering = false; 
        switchType(currentAttType); 
      }).withFailureHandler(function(err) {
        alert("送出失敗：" + err.message);
        if (btn) { btn.disabled = false; btn.innerHTML = originalText; }
        attIsRendering = false;
      }).saveAttendance(dateText, presentList, currentAttType, maleCount, femaleCount);
  }

  function filterAttList() {
    var searchInput = document.getElementById('attSearchInput');
    if (!searchInput) return;
    var kw = searchInput.value.trim().toLowerCase();
    var items = Array.prototype.slice.call(document.querySelectorAll('label.att-item'));
    var scrollArea = document.querySelector('.attendance-scroll-area');
    var wasFiltered = items.some(function(item) { return item.style.display === 'none'; });
    var anchor = null;
    if (kw === "" && wasFiltered && scrollArea && window.ListScrollAnchor) {
      anchor = window.ListScrollAnchor.capture(scrollArea, 'label.att-item');
    }

    items.forEach(function(item) {
      var nameEl = item.querySelector('.att-name');
      if (nameEl) {
          var name = nameEl.innerText.trim().toLowerCase();
          item.style.display = (kw === "" || name.includes(kw)) ? 'flex' : 'none';
      }
    });

    if (anchor && window.ListScrollAnchor) {
      var restoreAnchor = function() {
        window.ListScrollAnchor.restore(scrollArea, 'label.att-item', anchor);
      };
      if (typeof requestAnimationFrame === 'function') requestAnimationFrame(restoreAnchor);
      else setTimeout(restoreAnchor, 0);
    }
  }

  function toggleScanner() {
    if (!attUserId) attUserId = localStorage.getItem('att_uid');
    var scannerUrl = "https://jirehwang.github.io/LKC1958_June_1.github.io/apps/qrcodescanner.github.io/?userId=" + attUserId;
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
        renderAttendanceList(list, nfMale, nfFemale);
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
