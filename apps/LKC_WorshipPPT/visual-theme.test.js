const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const read = file => fs.readFileSync(path.join(__dirname, file), 'utf8');
const html = read('index.html');
const app = read('app.js');
const layoutGroups = read('layout-groups.js');
const styles = read('style.css');
const theme = read('theme.css');

test('preserves the existing three-column workspace layout', () => {
  assert.match(styles, /grid-template-columns:230px minmax\(430px,1fr\) minmax\(420px,560px\)/);
  assert.match(html, /<main class="workspace">[\s\S]*?<aside class="flow-panel">[\s\S]*?<section class="editor-panel">[\s\S]*?<aside class="preview-panel">/);
});

test('uses a calm liturgical palette without external visual assets', () => {
  assert.match(theme, /--accent:#243b63/);
  assert.match(theme, /--gold:#a88b50/);
  assert.match(theme, /--surface:#fffdf8/);
  assert.doesNotMatch(theme, /url\(['"]?https?:\/\//);
});

test('announces operation state and blocks editing behind a New Family-style swaying Bible loader', () => {
  assert.match(html, /id="save-state"[^>]*role="status"[^>]*aria-live="polite"[^>]*data-state="idle"/);
  assert.match(app, /target\.dataset\.state=state/);
  assert.match(app, /document\.body\.classList\.toggle\('is-busy',state==='busy'\)/);
  assert.match(theme, /@keyframes bible-sway/);
  assert.match(theme, /body\.is-busy::before/);
  assert.match(theme, /content:"\\1F4D6"/);
  assert.doesNotMatch(theme, /content:"正在處理"/);
  assert.match(theme, /font-size:28px/);
  assert.doesNotMatch(theme, /data:image\/svg\+xml/);
  assert.match(theme, /body\.is-busy::after\{content:"";position:fixed;z-index:110;inset:0/);
  assert.match(theme, /background:rgba\(255,253,248,\.58\)/);
  assert.match(theme, /pointer-events:auto/);
  assert.match(theme, /body\.is-busy\{cursor:progress\}/);
  assert.match(theme, /@media\(prefers-reduced-motion:reduce\)/);
});

test('reports the major shared-layout operations before awaiting Firebase', () => {
  assert.match(layoutGroups, /正在儲存輸出比例/);
  assert.match(layoutGroups, /正在儲存全教會共用版面群組/);
  assert.match(layoutGroups, /正在載入全教會共用版面配置/);
  assert.match(layoutGroups, /正在驗證版面設定密碼/);
});
