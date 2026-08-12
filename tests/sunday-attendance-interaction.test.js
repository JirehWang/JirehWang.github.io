const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const attendancePath = path.join(
  __dirname,
  '..',
  'apps',
  'LKC_SundayserviceAttendance',
  'attendance.js'
);
const attendanceSource = fs.readFileSync(attendancePath, 'utf8');

function makeClassList(owner, values = new Set()) {
  return {
    add(...names) { names.forEach(name => values.add(name)); owner.syncClassName(); },
    remove(...names) { names.forEach(name => values.delete(name)); owner.syncClassName(); },
    contains(name) { return values.has(name); },
    toggle(name, force) {
      const next = force === undefined ? !values.has(name) : force;
      if (next) values.add(name); else values.delete(name);
      owner.syncClassName();
      return next;
    },
    replaceFromClassName(value) {
      values.clear();
      String(value || '').split(/\s+/).filter(Boolean).forEach(name => values.add(name));
    }
  };
}

function makeCard(name = '王小明') {
  const card = {
    style: {},
    onclick: null,
    innerHTML: '',
    syncClassName() {
      this._className = Array.from(this.classListValues()).join(' ');
    },
    classListValues() {
      return this._classValues;
    },
    querySelector(selector) {
      if (selector === 'input') return checkbox;
      if (selector === '.att-name') return nameElement;
      return null;
    },
    querySelectorAll() { return []; }
  };
  card._classValues = new Set(['att-item', 'shadow-sm']);
  card.classList = makeClassList(card, card._classValues);
  card.classListValues = () => card._classValues;
  card.syncClassName = () => { card._className = Array.from(card._classValues).join(' '); };
  Object.defineProperty(card, 'className', {
    get() { return card._className; },
    set(value) {
      card._classValues.clear();
      String(value || '').split(/\s+/).filter(Boolean).forEach(name => card._classValues.add(name));
      card._className = String(value || '');
    }
  });
  const nameElement = { innerHTML: name };
  const checkbox = {
    checked: false,
    disabled: false,
    value: name,
    dataset: { uid: 'LK100' },
    parentElement: card
  };
  return { card, checkbox };
}

function makeAttendanceContext(card, checkbox, response) {
  const dateInput = { value: '2026-08-12', addEventListener() {} };
  const presentCount = { innerText: '0' };
  const container = {
    querySelector(selector) {
      return selector === 'input[data-uid="LK100"]' ? checkbox : null;
    },
    querySelectorAll() { return []; }
  };
  let successHandler = null;
  const run = {
    withSuccessHandler(handler) { successHandler = handler; return run; },
    getQuickSyncData() {
      successHandler(response);
    }
  };
  const context = {
    document: {
      addEventListener() {},
      getElementById(id) {
        if (id === 'attendanceDateInput') return dateInput;
        if (id === 'attendanceListBody') return container;
        if (id === 'presentCount') return presentCount;
        return null;
      }
    },
    google: { script: { run } },
    window: {},
    localStorage: {
      getItem() { return null; },
      setItem() {},
      removeItem() {}
    },
    formatDateToSlash(value) { return value; },
    Date,
    Number,
    String,
    Object,
    Boolean,
    Promise,
    setTimeout,
    clearTimeout,
    console,
    attIsRendering: false,
    remoteStatusSequence: 0,
    currentAttType: '主日',
    attUserId: 'device-a',
    localPendingActions: {},
    realtimeAttendanceTempReady: false,
    realtimeAttendanceTempEntries: {},
    lastClickTime: 0,
    DOUBLE_CLICK_DELAY: 350,
    confirmRevoke() {}
  };
  vm.createContext(context);
  const helperStart = attendanceSource.indexOf('function applyPendingSourceClass');
  const helperEnd = attendanceSource.indexOf('function startAutoSync', helperStart);
  vm.runInContext(attendanceSource.slice(helperStart, helperEnd), context, { filename: attendancePath });
  return context;
}

test('remote status refresh keeps the pending source class after card state updates', () => {
  const { card, checkbox } = makeCard();
  const context = makeAttendanceContext(card, checkbox, {
    activeList: [{ id: 'LK100', uid: 'LK100', name: '王小明', isChecked: true, isSubmitted: false, pendingSource: 'manual' }],
    nfMale: 0,
    nfFemale: 0
  });

  context.fetchRemoteStatus();

  assert.equal(card.classList.contains('pending-manual'), true);
  assert.equal(card.style.pointerEvents, 'auto');
  assert.equal(typeof card.onclick, 'function');
});

test('GAS polling does not clear a Firebase-owned source badge', () => {
  const { card, checkbox } = makeCard();
  card.classList.add('selected', 'pending-manual');
  const context = makeAttendanceContext(card, checkbox, {
    activeList: [{ id: 'LK100', name: '王小明', isChecked: false, isSubmitted: false }],
    nfMale: 0,
    nfFemale: 0
  });
  context.realtimeAttendanceTempReady = true;
  context.realtimeAttendanceTempEntries = { LK100: { checked: true, source: 'manual' } };

  context.fetchRemoteStatus();

  assert.equal(card.classList.contains('pending-manual'), true);
  assert.equal(card.classList.contains('selected'), true);
});

test('manual card toggles update the source badge immediately', () => {
  const { card, checkbox } = makeCard();
  const context = makeAttendanceContext(card, checkbox, { activeList: [] });
  context.enqueueAttendanceTemp = () => ({ uid: 'LK100' });
  context.flushAttendanceTempQueue = () => Promise.resolve();

  const toggleStart = attendanceSource.indexOf('function toggleCardStyle');
  const toggleEnd = attendanceSource.indexOf('function openAttendanceAddModal', toggleStart);
  vm.runInContext(attendanceSource.slice(toggleStart, toggleEnd), context, { filename: attendancePath });

  checkbox.checked = true;
  context.toggleCardStyle(checkbox);
  assert.equal(card.classList.contains('selected'), true);
  assert.equal(card.classList.contains('pending-manual'), true);

  checkbox.checked = false;
  context.toggleCardStyle(checkbox);
  assert.equal(card.classList.contains('pending-manual'), false);
});

test('attendance card clicks remain bound to one local toggle action', () => {
  const { card, checkbox } = makeCard();
  const context = makeAttendanceContext(card, checkbox, { activeList: [] });
  let toggleCount = 0;
  context.toggleCardStyle = () => { toggleCount += 1; };

  context.bindAttendanceCardClick(card);
  card.onclick({ preventDefault() {} });

  assert.equal(checkbox.checked, true);
  assert.equal(toggleCount, 1);
});
