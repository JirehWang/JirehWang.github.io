// 全域變數
let events = [];
let eventIdCounter = 1;
let ministryIdCounter = 1; // 事工細項計數器
let sermonIdCounter = 1;   // 講道資訊計數器

// ====================
// 🚀 系統啟動與安全連線
// ====================

// 🛡️ API 哨兵：確保 config.js 已經準備好
async function ensureAPIReady() {
    let retryCount = 0;
    while (typeof window.churchAPI !== 'function' && retryCount < 50) {
        await new Promise(resolve => setTimeout(resolve, 100)); 
        retryCount++;
    }
    if (typeof window.churchAPI !== 'function') {
        throw new Error("安全路由載入逾時，請檢查 config.js。");
    }
}

// 初始化
document.addEventListener('DOMContentLoaded', async function() {
    try {
        // 第一時間顯示載入中
        showLoading("🚀 正在啟動行事曆系統...");
        
        // 等待路由就緒
        await ensureAPIReady();
        
        loadFromLocalStorage();
        renderEvents();
        console.log("✅ 行事曆系統安全啟動");
    } catch (e) {
        alert("系統啟動失敗：" + e.message);
    } finally {
        hideLoading();
    }
});

// 核心連線函式 (透過中央路由，完美隱藏真實 URL 與 Token)
async function callCloudAPI(action, data = {}) {
    if (typeof window.churchAPI !== 'function') throw new Error("API 尚未就緒");
    return await window.churchAPI(action, data);
}

// --- Loading 控制 ---
function showLoading(msg = "處理中...") {
    const textEl = document.getElementById('overlay-text');
    const overlayEl = document.getElementById('loading-overlay');
    if (textEl) textEl.innerText = msg;
    if (overlayEl) overlayEl.style.display = 'flex';
}
function hideLoading() {
    const overlayEl = document.getElementById('loading-overlay');
    if (overlayEl) overlayEl.style.display = 'none';
}

// ====================
// 聚會新增與基本管理
// ====================

// 新增單一對話框的專屬邏輯
function addEvent() {
    document.getElementById('singleEventModal').classList.remove('hidden');
    document.getElementById('singleDate').value = new Date().toISOString().split('T')[0];
    document.getElementById('singleName').value = '';
    document.getElementById('singleCategory').value = '台華語聚會';
}

function closeSingleEventModal() {
    document.getElementById('singleEventModal').classList.add('hidden');
}

function confirmSingleEventAdd() {
    const dateStr = document.getElementById('singleDate').value;
    const name = document.getElementById('singleName').value.trim();
    const category = document.getElementById('singleCategory').value;

    if (!dateStr) {
        alert('請選擇日期');
        return;
    }

    const event = {
        id: eventIdCounter++,
        date: dateStr,
        name: name,
        category: category,
        ministryItems: [], 
        sermons: [],       
        showSub: true      
    };
    
    events.push(event);
    sortEvents(); // 新增後自動排序
    
    renderEvents();
    saveToLocalStorage();
    closeSingleEventModal();
}

function deleteEvent(eventId) {
    if (confirm('確定要刪除此聚會嗎？')) {
        events = events.filter(e => e.id !== eventId);
        renderEvents();
        saveToLocalStorage();
    }
}

// 更新聚會資料 (加入自動排序邏輯)
function updateEvent(eventId, field, value) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        event[field] = value;
        // 如果改的是日期，就觸發重新排序與渲染
        if (field === 'date') {
            sortEvents();
            renderEvents();
        }
        saveToLocalStorage();
    }
}

function toggleSubItems(eventId) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        event.showSub = !event.showSub;
        renderEvents(); 
    }
}

// ====================
// 排序與手動調整順序功能
// ====================
function sortEvents() {
    // 依據日期由舊到新排序 (Stable Sort，若日期相同則維持目前的手動順序)
    events.sort((a, b) => new Date(a.date) - new Date(b.date));
}

function moveEventUp(eventId) {
    const index = events.findIndex(e => e.id === eventId);
    if (index > 0) {
        // 限制只能與「相同日期」的聚會交換順序
        if (events[index].date !== events[index - 1].date) {
            alert('💡 只能與「相同日期」的聚會互相調整順序喔！\n若要跨日移動，請直接修改日期欄位。');
            return;
        }
        // 陣列元素交換
        const temp = events[index];
        events[index] = events[index - 1];
        events[index - 1] = temp;
        
        saveToLocalStorage();
        renderEvents();
    }
}

function moveEventDown(eventId) {
    const index = events.findIndex(e => e.id === eventId);
    if (index < events.length - 1) {
        // 限制只能與「相同日期」的聚會交換順序
        if (events[index].date !== events[index + 1].date) {
            alert('💡 只能與「相同日期」的聚會互相調整順序喔！\n若要跨日移動，請直接修改日期欄位。');
            return;
        }
        // 陣列元素交換
        const temp = events[index];
        events[index] = events[index + 1];
        events[index + 1] = temp;
        
        saveToLocalStorage();
        renderEvents();
    }
}

// ====================
// 拖曳排序功能變數與邏輯
// ====================
let draggedEventId = null;

function handleDragStart(e, eventId) {
    draggedEventId = eventId;
    // 稍微延遲加上 class，讓拖曳出去的影子保持原樣，但留在原地的卡片變半透明
    setTimeout(() => e.target.classList.add('dragging'), 0);
    e.dataTransfer.effectAllowed = 'move';
}

function handleDragOver(e) {
    e.preventDefault(); // 必須 preventDefault 才能允許 drop
    const card = e.target.closest('.event-card');
    if (card && !card.classList.contains('dragging')) {
        card.classList.add('drag-over');
    }
}

function handleDragLeave(e) {
    const card = e.target.closest('.event-card');
    if (card) {
        card.classList.remove('drag-over');
    }
}

function handleDrop(e, targetEventId) {
    e.preventDefault();
    
    // 清除所有視覺效果
    document.querySelectorAll('.event-card').forEach(c => {
        c.classList.remove('dragging');
        c.classList.remove('drag-over');
    });

    if (!draggedEventId || draggedEventId === targetEventId) return;

    const draggedIndex = events.findIndex(ev => ev.id === draggedEventId);
    const targetIndex = events.findIndex(ev => ev.id === targetEventId);

    if (draggedIndex === -1 || targetIndex === -1) return;

    // 防呆：檢查是否為同一天
    if (events[draggedIndex].date !== events[targetIndex].date) {
        alert('💡 只能與「相同日期」的聚會互相調整順序喔！\n若要跨日移動，請直接修改日期欄位。');
        return;
    }

    // 執行陣列重新排列 (計算相對位置)
    const targetItem = events[targetIndex];
    const [draggedItem] = events.splice(draggedIndex, 1);
    const newTargetIndex = events.indexOf(targetItem);
    
    // 判斷是往下拖還是往上拖，決定插入在目標的前面還是後面
    if (draggedIndex < targetIndex) {
        events.splice(newTargetIndex + 1, 0, draggedItem);
    } else {
        events.splice(newTargetIndex, 0, draggedItem);
    }
    
    saveToLocalStorage();
    renderEvents();
    draggedEventId = null;
}

// ====================
// 事工細項相關功能
// ====================
function addMinistryItem(eventId) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        event.ministryItems.push({
            id: ministryIdCounter++,
            date: new Date().toISOString().split('T')[0],
            content: ''
        });
        event.showSub = true;
        renderEvents();
        saveToLocalStorage();
    }
}

function deleteMinistryItem(eventId, minId) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        event.ministryItems = event.ministryItems.filter(m => m.id !== minId);
        renderEvents();
        saveToLocalStorage();
    }
}

function updateMinistryItem(eventId, minId, field, value) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        const item = event.ministryItems.find(m => m.id === minId);
        if (item) {
            item[field] = value;
            saveToLocalStorage();
        }
    }
}

// ====================
// 講道資訊相關功能
// ====================
function addSermon(eventId, type) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        event.sermons.push({
            id: sermonIdCounter++,
            type: type, 
            title: '',
            speaker: '',
            scripture: '',
            callToWorship: '',
            goldenVerse: '',
            hymns: '',
            description: ''
        });
        event.showSub = true;
        renderEvents();
        saveToLocalStorage();
    }
}

function deleteSermon(eventId, sermonId) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        event.sermons = event.sermons.filter(s => s.id !== sermonId);
        renderEvents();
        saveToLocalStorage();
    }
}

function updateSermon(eventId, sermonId, field, value) {
    const event = events.find(e => e.id === eventId);
    if (event) {
        const sermon = event.sermons.find(s => s.id === sermonId);
        if (sermon) {
            sermon[field] = value;
            saveToLocalStorage();
        }
    }
}

// ====================
// 核心：取得「過濾後」的資料
// ====================
function getFilteredEvents() {
    const startInput = document.getElementById('searchStartDate');
    const endInput = document.getElementById('searchEndDate');
    
    if (!startInput || !endInput) return events;
    
    const start = startInput.value;
    const end = endInput.value;
    
    if (!start && !end) return events;

    return events.filter(event => {
        const eventDate = event.date;
        if (start && end) return eventDate >= start && eventDate <= end;
        if (start) return eventDate >= start;
        if (end) return eventDate <= end;
        return true;
    });
}

function filterEvents() {
    renderEvents(); 
}

function clearFilter() {
    document.getElementById('searchStartDate').value = '';
    document.getElementById('searchEndDate').value = '';
    renderEvents();
}

// ====================
// 渲染所有聚會
// ====================
function renderEvents() {
    const container = document.getElementById('eventList');
    if (!container) return;
    container.innerHTML = '';

    const dataToRender = getFilteredEvents();

    if (dataToRender.length === 0) {
        const hasSearch = document.getElementById('searchStartDate')?.value || document.getElementById('searchEndDate')?.value;
        if (hasSearch) {
            container.innerHTML = '<p class="loading">找不到符合此日期區間的聚會資料</p>';
        } else {
            container.innerHTML = '<p class="loading">尚無聚會資料，請點擊「新增聚會」開始建立</p>';
        }
        return;
    }

    dataToRender.forEach(event => {
        const ministryItems = event.ministryItems || [];
        const sermons = event.sermons || [];

        const card = document.createElement('div');
        card.className = 'event-card';
        
        // --- 綁定拖曳相關事件 ---
        card.setAttribute('draggable', 'true');
        card.setAttribute('ondragstart', `handleDragStart(event, ${event.id})`);
        card.setAttribute('ondragover', `handleDragOver(event)`);
        card.setAttribute('ondragleave', `handleDragLeave(event)`);
        card.setAttribute('ondrop', `handleDrop(event, ${event.id})`);

        card.innerHTML = `
            <div class="event-header" style="display: flex; flex-wrap: wrap; gap: 15px; align-items: flex-end;">
                
                <div style="cursor: grab; font-size: 20px; color: #a0aec0; padding-right: 5px; padding-bottom: 5px; display: flex; align-items: center;" title="按住拖曳以調整順序">
                    ☰
                </div>

                <div class="input-group" style="margin-bottom: 0;">
                    <label>日期</label>
                    <input type="date" value="${event.date}" 
                           onchange="updateEvent(${event.id}, 'date', this.value)">
                </div>
                <div class="input-group" style="margin-bottom: 0;">
                    <label>聚會名稱</label>
                    <input type="text" placeholder="輸入聚會名稱" value="${event.name}"
                           onchange="updateEvent(${event.id}, 'name', this.value)">
                </div>
                <div class="input-group" style="margin-bottom: 0;">
                    <label>聚會類別</label>
                    <select onchange="updateEvent(${event.id}, 'category', this.value)">
                        <option ${event.category === '台華語聚會' ? 'selected' : ''}>台華語聚會</option>
                        <option ${event.category === '聯合聚會' ? 'selected' : ''}>聯合聚會</option>
                    </select>
                </div>
                
                <div style="display: flex; gap: 8px; margin-left: auto;">
                    <button class="btn btn-small" style="padding: 4px 10px; background: #e2e8f0; color: #2d3748; border: 1px solid #cbd5e0;" onclick="moveEventUp(${event.id})" title="與同日期的聚會對調(上移)">⬆️</button>
                    <button class="btn btn-small" style="padding: 4px 10px; background: #e2e8f0; color: #2d3748; border: 1px solid #cbd5e0;" onclick="moveEventDown(${event.id})" title="與同日期的聚會對調(下移)">⬇️</button>
                    <button class="btn btn-danger btn-small" onclick="deleteEvent(${event.id})">刪除</button>
                </div>
            </div>
            
            <button class="btn toggle-sub btn-small" onclick="toggleSubItems(${event.id})" style="margin-top: 15px;">
                ${event.showSub ? '▼' : '▶'} 聚會詳情 (事工:${ministryItems.length} / 講道:${sermons.length})
            </button>
            
            <div class="sub-items ${event.showSub ? '' : 'hidden'}" style="display: flex; flex-direction: column; gap: 20px;">
                
                <div style="background: #fffafa; border: 1px solid #fed7d7; padding: 15px; border-radius: 8px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; border-bottom: 1px solid #feb2b2; padding-bottom: 10px;">
                        <h4 style="margin: 0; color: #c53030;">📋 事工細項</h4>
                        <button class="btn btn-primary btn-small" onclick="addMinistryItem(${event.id})" style="background: #c53030; border-color: #c53030;">➕ 新增事工</button>
                    </div>
                    ${ministryItems.map(min => `
                        <div class="sub-item" style="display: flex; gap: 15px; align-items: flex-end; margin-bottom: 10px; padding: 10px; background: white; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div class="input-group" style="flex: 1; margin-bottom: 0;">
                                <label>籌備日期</label>
                                <input type="date" value="${min.date || ''}" 
                                       onchange="updateMinistryItem(${event.id}, ${min.id}, 'date', this.value)">
                            </div>
                            <div class="input-group" style="flex: 2; margin-bottom: 0;">
                                <label>事工內容</label>
                                <input type="text" placeholder="開會討論、預演準備..." value="${min.content || ''}" 
                                       onchange="updateMinistryItem(${event.id}, ${min.id}, 'content', this.value)">
                            </div>
                            <button class="btn btn-danger btn-small" style="height: 38px;" 
                                    onclick="deleteMinistryItem(${event.id}, ${min.id})">刪除</button>
                        </div>
                    `).join('')}
                </div>

                <div style="background: #f0f4f8; border: 1px solid #bee3f8; padding: 15px; border-radius: 8px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; border-bottom: 1px solid #90cdf4; padding-bottom: 10px;">
                        <h4 style="margin: 0; color: #2b6cb0;">🎙️ 講道資訊</h4>
                        <div style="display: flex; gap: 10px;">
                            <button class="btn btn-primary btn-small" onclick="addSermon(${event.id}, '台語/聯合')" style="background: #3182ce; border-color: #3182ce;">➕ 台語/聯合講道</button>
                            <button class="btn btn-primary btn-small" onclick="addSermon(${event.id}, '華語')" style="background: #38a169; border-color: #38a169;">➕ 華語講道</button>
                        </div>
                    </div>
                    ${sermons.map(sermon => {
                        const isMandarin = sermon.type === '華語';
                        const hideStyle = isMandarin ? 'display: none;' : '';
                        const tagStyle = isMandarin ? 'background: #c6f6d5; color: #22543d;' : 'background: #bee3f8; color: #2a4365;';
                        
                        return `
                        <div class="sub-item" style="border-left: 4px solid ${isMandarin ? '#38a169' : '#3182ce'}; padding: 15px; background: white; margin-bottom: 15px; border-radius: 0 6px 6px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                                <span style="padding: 4px 10px; border-radius: 4px; font-size: 13px; font-weight: bold; ${tagStyle}">${sermon.type}</span>
                                <button class="btn btn-danger btn-small" onclick="deleteSermon(${event.id}, ${sermon.id})">刪除此講道</button>
                            </div>
                            
                            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 10px;">
                                <div class="input-group" style="margin-bottom: 0;">
                                    <label>講題</label>
                                    <input type="text" value="${sermon.title || ''}" onchange="updateSermon(${event.id}, ${sermon.id}, 'title', this.value)">
                                </div>
                                <div class="input-group" style="margin-bottom: 0;">
                                    <label>講員</label>
                                    <input type="text" value="${sermon.speaker || ''}" onchange="updateSermon(${event.id}, ${sermon.id}, 'speaker', this.value)">
                                </div>
                                <div class="input-group" style="margin-bottom: 0;">
                                    <label>經文</label>
                                    <input type="text" value="${sermon.scripture || ''}" onchange="updateSermon(${event.id}, ${sermon.id}, 'scripture', this.value)">
                                </div>
                                <div class="input-group" style="margin-bottom: 0;">
                                    <label>宣召</label>
                                    <input type="text" value="${sermon.callToWorship || ''}" onchange="updateSermon(${event.id}, ${sermon.id}, 'callToWorship', this.value)">
                                </div>
                                <div class="input-group" style="margin-bottom: 0; ${hideStyle}">
                                    <label>金句</label>
                                    <input type="text" value="${sermon.goldenVerse || ''}" onchange="updateSermon(${event.id}, ${sermon.id}, 'goldenVerse', this.value)">
                                </div>
                                <div class="input-group" style="margin-bottom: 0; ${hideStyle}">
                                    <label>詩歌/聖詩</label>
                                    <input type="text" value="${sermon.hymns || ''}" onchange="updateSermon(${event.id}, ${sermon.id}, 'hymns', this.value)">
                                </div>
                            </div>

                            <div class="input-group" style="margin-top: 15px; margin-bottom: 0;">
                                <label>內容描述 / 備註</label>
                                <textarea placeholder="詳細說明..." style="min-height: 60px;"
                                          onchange="updateSermon(${event.id}, ${sermon.id}, 'description', this.value)">${sermon.description || ''}</textarea>
                            </div>
                        </div>
                    `}).join('')}
                </div>
                
            </div>
        `;
        container.appendChild(card);
    });
}

// ====================
// 儲存與匯出功能
// ====================

function saveToLocalStorage() {
    localStorage.setItem('churchEvents', JSON.stringify(events));
}

function loadFromLocalStorage() {
    const saved = localStorage.getItem('churchEvents');
    if (saved) {
        events = JSON.parse(saved);
        sortEvents(); // 載入時確保排序正確
        recalculateCounters();
    }
}

function recalculateCounters() {
    eventIdCounter = Math.max(...events.map(e => e.id), 0) + 1;
    
    events.forEach(e => {
        if (!e.ministryItems) e.ministryItems = [];
        if (!e.sermons) e.sermons = [];
    });

    const allMinIds = events.flatMap(e => e.ministryItems.map(m => m.id));
    ministryIdCounter = Math.max(...allMinIds, 0) + 1;
    
    const allSermonIds = events.flatMap(e => e.sermons.map(s => s.id));
    sermonIdCounter = Math.max(...allSermonIds, 0) + 1;
}

function exportToExcel() {
    if (events.length === 0) {
        alert('沒有資料可以匯出');
        return;
    }

    let csv = '\uFEFF';
    csv += '日期,聚會名稱,聚會類別,分類,細項日期,內容/講題,講員,經文,宣召,金句,詩歌,備註\n';

    events.forEach(event => {
        const hasMinistry = event.ministryItems && event.ministryItems.length > 0;
        const hasSermons = event.sermons && event.sermons.length > 0;

        if (!hasMinistry && !hasSermons) {
            csv += `${event.date},"${event.name}","${event.category}",,,,,,,,,\n`;
        } else {
            if (hasMinistry) {
                event.ministryItems.forEach(min => {
                    csv += `${event.date},"${event.name}","${event.category}","籌備事工","${min.date || ''}","${min.content || ''}",,,,,,,\n`;
                });
            }
            if (hasSermons) {
                event.sermons.forEach(sermon => {
                    csv += `${event.date},"${event.name}","${event.category}","講道(${sermon.type})",,"${sermon.title || ''}","${sermon.speaker || ''}","${sermon.scripture || ''}","${sermon.callToWorship || ''}","${sermon.goldenVerse || ''}","${sermon.hymns || ''}","${(sermon.description || '').replace(/\n/g, ' ')}"\n`;
                });
            }
        }
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `教會行事曆_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ====================
// 🌟 雲端同步 (使用中央路由)
// ====================

async function saveToGAS() {
    const saveBtn = document.querySelector('button[onclick="saveToGAS()"]');
    const originalText = saveBtn.innerHTML;
    saveBtn.innerHTML = '⏳ 儲存中...';
    saveBtn.disabled = true;

    try {
        const result = await callCloudAPI('save', events);
        
        if (result.success) {
            alert('✅ 資料已成功儲存到雲端！');
        } else {
            throw new Error(result.error || '後端回傳失敗');
        }
    } catch (error) {
        console.error('儲存失敗細節:', error);
        alert(`❌ 儲存失敗！\n原因: ${error.message}`);
    } finally {
        saveBtn.innerHTML = originalText;
        saveBtn.disabled = false;
    }
}

async function loadFromGAS() {
    const loadBtn = document.querySelector('button[onclick="loadFromGAS()"]');
    const originalText = loadBtn.innerHTML;
    loadBtn.innerHTML = '⏳ 載入中...';
    loadBtn.disabled = true;

    try {
        // 🌟 改用 POST 方式讀取，完美避開 GET 跨網域問題
        const data = await callCloudAPI('load');
        
        if (data.success && data.events) {
            events = data.events;
            sortEvents(); // 從雲端載入後確保排序正確
            recalculateCounters(); 
            renderEvents();
            saveToLocalStorage(); 
            alert('✅ 資料已從雲端載入！');
        }
    } catch (error) {
        console.error('載入失敗:', error);
        alert('載入失敗，請檢查網路連線');
    } finally {
        loadBtn.innerHTML = originalText;
        loadBtn.disabled = false;
    }
}

// ====================
// 批量新增功能
// ====================
let selectedWeekdays = new Set();
let previewDates = [];

function openBatchModal() {
    document.getElementById('batchModal').classList.remove('hidden');
    document.getElementById('batchStartDate').value = '';
    document.getElementById('batchEndDate').value = '';
    document.getElementById('batchName').value = '';
    document.getElementById('batchCategory').value = '台華語聚會';
    selectedWeekdays.clear();
    previewDates = [];
    
    document.querySelectorAll('.weekday-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    document.getElementById('datePreview').innerHTML = '請選擇日期區間和星期';
    document.getElementById('datePreview').classList.remove('has-dates');
}

function closeBatchModal() {
    document.getElementById('batchModal').classList.add('hidden');
}

function toggleWeekday(day) {
    if (selectedWeekdays.has(day)) {
        selectedWeekdays.delete(day);
    } else {
        selectedWeekdays.add(day);
    }
    
    const btn = document.querySelector(`.weekday-btn[data-day="${day}"]`);
    btn.classList.toggle('active');
}

function previewBatchDates() {
    const startDate = document.getElementById('batchStartDate').value;
    const endDate = document.getElementById('batchEndDate').value;
    
    if (!startDate || !endDate) return;
    if (selectedWeekdays.size === 0) return;
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    if (start > end) {
        alert('開始日期不能晚於結束日期');
        return;
    }
    
    previewDates = [];
    const current = new Date(start);
    
    while (current <= end) {
        const dayOfWeek = current.getDay();
        if (selectedWeekdays.has(dayOfWeek)) {
            previewDates.push(new Date(current));
        }
        current.setDate(current.getDate() + 1);
    }
    
    const previewDiv = document.getElementById('datePreview');
    if (previewDates.length === 0) {
        previewDiv.innerHTML = '在選定的日期區間內，沒有符合條件的日期';
        previewDiv.classList.remove('has-dates');
    } else {
        const weekdayNames = ['日', '一', '二', '三', '四', '五', '六'];
        let html = `<div class="preview-count">共 ${previewDates.length} 個日期</div>`;
        
        previewDates.forEach(date => {
            const dateStr = date.toISOString().split('T')[0];
            const dayName = weekdayNames[date.getDay()];
            html += `<span class="preview-date-item">${dateStr} (${dayName})</span>`;
        });
        
        previewDiv.innerHTML = html;
        previewDiv.classList.add('has-dates');
    }
}

function confirmBatchAdd() {
    const name = document.getElementById('batchName').value.trim();
    const category = document.getElementById('batchCategory').value;
    
    if (previewDates.length === 0) {
        alert('請先預覽日期，確認有可新增的日期');
        return;
    }
    
    let addedCount = 0;
    previewDates.forEach(date => {
        const dateStr = date.toISOString().split('T')[0];
        const exists = events.some(e => e.date === dateStr && e.name === name);
        
        if (!exists) {
            events.push({
                id: eventIdCounter++,
                date: dateStr,
                name: name,
                category: category,
                ministryItems: [],
                sermons: [],
                showSub: false
            });
            addedCount++;
        }
    });
    
    sortEvents();
    renderEvents();
    saveToLocalStorage();
    closeBatchModal();
    
    alert(`成功新增 ${addedCount} 個聚會！${previewDates.length - addedCount > 0 ? '\n(' + (previewDates.length - addedCount) + ' 個重複的聚會已略過)' : ''}`);
}

// ====================
// 🤖 AI 批量匯入講道功能
// ====================
let parsedAiData = [];

function updateAiPrompt() {
    const category = document.getElementById('aiCategory').value;
    const promptEl = document.getElementById('aiPrompt');
    
    let basePrompt = `你是一個教會行政助理。請將以下牧師提供的講道資訊，轉換為嚴格的 JSON 格式陣列。
請自動判斷年份，若無標示年份請預設為今年。
請依照以下格式輸出（只輸出 JSON，不要包含 \`\`\`json 標籤或其他說明文字）：
[
  {
    "date": "YYYY-MM-DD",
    "eventName": "特別聚會名稱 (如: 復活節聯合禮拜、母親節，若無則留空)",
    "speaker": "講員 (若無則留空)",
    "title": "講題 (去掉中括號【】，若無則留空)",
    "scripture": "經文內容，若無則留空",
    "callToWorship": "宣召內容，若無則留空"`;

    if (category === '台語/聯合') {
        basePrompt += `,\n    "goldenVerse": "金句內容",\n    "hymns": "聖詩/回應詩內容"\n  }\n]`;
    } else {
        basePrompt += `\n  }\n]`;
    }
    
    promptEl.value = basePrompt;
}

function openAiImportModal() {
    document.getElementById('aiImportModal').classList.remove('hidden');
    document.getElementById('aiCategory').value = '台語/聯合';
    updateAiPrompt(); 
    document.getElementById('aiRawText').value = '';
    document.getElementById('aiPreview').innerHTML = '請輸入文字並點擊「開始解析」';
    document.getElementById('aiPreview').classList.remove('has-dates');
    document.getElementById('btnConfirmAi').disabled = true;
    parsedAiData = [];
}

function closeAiImportModal() {
    document.getElementById('aiImportModal').classList.add('hidden');
}

function closeModalOnBackdrop(event) {
    if (event.target.id === 'batchModal') {
        closeBatchModal();
    } else if (event.target.id === 'singleEventModal') {
        closeSingleEventModal();
    } else if (event.target.id === 'aiImportModal') {
        closeAiImportModal();
    }
}

async function processAiText() {
    const rawText = document.getElementById('aiRawText').value.trim();
    const prompt = document.getElementById('aiPrompt').value.trim();
    const previewDiv = document.getElementById('aiPreview');
    const confirmBtn = document.getElementById('btnConfirmAi');

    if (!rawText) {
        alert('請先貼上牧師的原始文字');
        return;
    }

    previewDiv.innerHTML = '<div class="loading">⏳ AI 正在努力解析中，請稍候...</div>';
    confirmBtn.disabled = true;

    try {
        // 🌟 改用中央路由發送 AI 解析請求
        const data = await callCloudAPI('ai_parse', { prompt: prompt, rawText: rawText });

        if (data.success) {
            parsedAiData = JSON.parse(data.result);

            if (parsedAiData.length > 0) {
                let html = `<div class="preview-count">成功解析出 ${parsedAiData.length} 筆講道資訊</div>`;
                parsedAiData.forEach(item => {
                    html += `
                        <div class="sub-item" style="margin-top: 8px;">
                            <strong>📅 ${item.date} ${item.eventName ? `(${item.eventName})` : ''}</strong><br>
                            <span style="color: #667eea; font-weight:bold;">${item.title || '(無標題)'}</span>
                            ${item.speaker ? `<span style="color: #e53e3e;"> - ${item.speaker}</span>` : ''}<br>
                            <div style="font-size: 13px; color: #718096; margin-top: 4px; line-height: 1.5;">
                                ${item.scripture ? `📖 經文：${item.scripture}<br>` : ''}
                                ${item.callToWorship ? `📣 宣召：${item.callToWorship}<br>` : ''}
                                ${item.goldenVerse ? `⭐ 金句：${item.goldenVerse}<br>` : ''}
                                ${item.hymns ? `🎵 詩歌：${item.hymns}` : ''}
                            </div>
                        </div>`;
                });
                previewDiv.innerHTML = html;
                previewDiv.classList.add('has-dates');
                confirmBtn.disabled = false;
            } else {
                previewDiv.innerHTML = '無法解析出有效資訊，請調整提示詞或檢查原始文字。';
            }
        } else {
            throw new Error(data.error || '後端回傳未知的錯誤');
        }

    } catch (error) {
        console.error('AI 解析錯誤:', error);
        previewDiv.innerHTML = `<span style="color: #e53e3e;">❌ 解析失敗，請確認資料格式或聯絡管理員。(${error.message})</span>`;
    }
}

function confirmAiImport() {
    if (parsedAiData.length === 0) return;

    const selectedLanguage = document.getElementById('aiCategory').value;
    let addedCount = 0;

    parsedAiData.forEach(aiItem => {
        let targetEvent = events.find(e => e.date === aiItem.date);
        
        let finalEventName = '主日崇拜';
        let eventCategory = '台華語聚會'; 

        if (aiItem.eventName) {
            finalEventName += ` (${aiItem.eventName})`;
            if (aiItem.eventName.includes('聯合')) {
                eventCategory = '聯合聚會';
            }
        }

        if (!targetEvent) {
            targetEvent = {
                id: eventIdCounter++,
                date: aiItem.date,
                name: finalEventName,
                category: eventCategory,
                ministryItems: [],
                sermons: [],
                showSub: true
            };
            events.push(targetEvent);
        } else {
            targetEvent.showSub = true; 
            if (aiItem.eventName && targetEvent.name === '主日崇拜') {
                targetEvent.name = finalEventName;
            }
            if (aiItem.eventName && aiItem.eventName.includes('聯合')) {
                targetEvent.category = '聯合聚會';
            }
        }

        targetEvent.sermons.push({
            id: sermonIdCounter++,
            type: selectedLanguage, 
            title: aiItem.title || '',
            speaker: aiItem.speaker || '',
            scripture: aiItem.scripture || '',
            callToWorship: aiItem.callToWorship || '',
            goldenVerse: aiItem.goldenVerse || '',
            hymns: aiItem.hymns || '',
            description: ''
        });
        
        addedCount++;
    });

    sortEvents();
    renderEvents();
    saveToLocalStorage();
    closeAiImportModal();

    alert(`✅ 成功將 ${addedCount} 筆講道資訊整合進行事曆！`);
}
