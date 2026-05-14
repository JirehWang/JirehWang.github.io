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
  
  let globalGroupConfig = {};

  function updateBadgeUI(sleepMode) {
    var badge = document.getElementById('presentCountBadge');
    var pCountEl = document.getElementById('presentCount');
    var count = pCountEl ? pCountEl.innerText : '0';
    if (!badge) return;
    if (sleepMode) {
      badge.className = "badge bg-secondary text-white p-2 small shadow-sm d-flex align-items-center flex-grow-1 justify-content-center";
      badge.innerHTML = "💤 休眠中(點擊) | 已出席：<span id='presentCount' class='mx-1'>" + count + "</span> 人";
    } else {
      badge.className = "badge bg-success text-white p-2 small shadow-sm d-flex align-items-center flex-grow-1 justify-content-center";
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

  var today = new Date();
  var dateDisplay = document.getElementById('todayDateDisplay');
  if (dateDisplay) dateDisplay.innerText = "📅 " + today.getFullYear() + "/" + (today.getMonth()+1) + "/" + today.getDate();

  function loadGroupConfig(targetCategory = null, targetGroup = null) {
    google.script.run.withSuccessHandler(config => {
      globalGroupConfig = config;
      renderCategorySelect(targetCategory);
      updateGroupSelect(targetGroup);
      startAutoSync();
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
      .getSmartAttendanceList(requestedType, attUserId); 
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
      if (localPendingActions[memberKey] && (now - localPendingActions[memberKey].time < 5000)) {
        isChecked = localPendingActions[memberKey].state;
        isLocked = false;
      }
      var checkState = isChecked ? "checked" : "";
      var isDisabled = (isSubmitted || isLocked) ? "disabled" : "";
      var statusColor = m.gender === '男' ? '#007bff' : (m.gender === '女' ? '#e64980' : '#6c757d');
      var lockIcon = ""; 
      if (isSubmitted) {
        label.classList.add('submitted');
        statusColor = '#198754';
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
        statusColor = '#e9ecef';
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
  }

  function toggleCardStyle(checkbox) {
    var isChecked = checkbox.checked;
    var uid = checkbox.dataset.uid || checkbox.value;  // 優先用 UID
    if (isChecked) checkbox.parentElement.classList.add('selected');
    else checkbox.parentElement.classList.remove('selected');
    localPendingActions[uid] = { time: Date.now(), state: isChecked };
    google.script.run.withFailureHandler(function(err) {
        checkbox.checked = !isChecked;
        checkbox.parentElement.classList.toggle('selected');
        delete localPendingActions[uid];
    }).syncClickToServer(uid, isChecked, currentAttType, attUserId);
  }

  function openAttendanceAddModal() {
    var modal = document.getElementById('attendanceAddModal');
    if (modal) modal.style.display = 'block';
    document.getElementById('editName_Att').value = "";
    document.getElementById('editGender_Att').value = "男";
    document.getElementById('editNote_Att').value = "";
    document.getElementById('editIsExcluded_Att').checked = false;
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
    google.script.run.withSuccessHandler(function(result) {
        var list = Array.isArray(result) ? result : (result.activeList || []);
        var nfMale = result.nfMale || 0;
        var nfFemale = result.nfFemale || 0;
        renderAttendanceList(list, nfMale, nfFemale);
        attIsRendering = false;
    }).getSmartAttendanceList(currentAttType, attUserId);
  }

  function confirmRevoke(uid, displayName) {
    if (navigator.vibrate) navigator.vibrate(50);
    if (confirm("確定要撤銷 [" + (displayName || uid) + "] 的送出紀錄嗎？")) { executeRevoke(uid, displayName); }
  }

function executeRevoke(uid, displayName) {
    var btn = document.getElementById('submitBtn');
    var originalText = "確認送出";
    if (btn) { btn.disabled = true; btn.innerHTML = '正在撤銷...'; }
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
    }).revokeAttendance(uid, currentAttType, attUserId);
}

  function fetchRemoteStatus() {
    var searchInput = document.getElementById('attSearchInput');
    if (attIsRendering || (searchInput && searchInput.value)) return;
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
    }).getQuickSyncData(currentAttType, attUserId);
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
    var dateDisplay = document.getElementById('todayDateDisplay');
    var dateText = dateDisplay ? dateDisplay.innerText.replace('📅 ','') : '';
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
    document.querySelectorAll('label.att-item').forEach(function(item) {
      var nameEl = item.querySelector('.att-name');
      if (nameEl) {
          var name = nameEl.innerText.trim().toLowerCase();
          item.style.display = (kw === "" || name.includes(kw)) ? 'flex' : 'none';
      }
    });
  }

  function toggleScanner() {
    if (!attUserId) attUserId = localStorage.getItem('att_uid');
    var scannerUrl = "https://jirehwang.github.io/qrcodescanner.github.io/?userId=" + attUserId;
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
    const finalUrl = "https://jirehwang.github.io/LKC_-SundayserviceAttendance.github.io/?cat=" + encodeURIComponent(cat) + "&grp=" + encodeURIComponent(grp) + "&v=" + timestamp;
    
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

  // 頁面載入時初始化
  loadGroupConfig();
