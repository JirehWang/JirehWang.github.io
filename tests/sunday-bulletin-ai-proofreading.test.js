const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const proofreader = require('../apps/LKC_SundayBulletin/js/reports-ai.js');

test('proofreading payload keeps field ids and excludes blank fields', () => {
  assert.deepEqual(
    proofreader.buildProofreadingPayload([
      { id: 'announcements.0', text: '  主日聚會提醒  ' },
      { id: 'announcements.1', text: '   ' },
      { id: 'prayer.homeRest', text: '請為病中的家人代禱' }
    ]),
    [
      { id: 'announcements.0', text: '主日聚會提醒' },
      { id: 'prayer.homeRest', text: '請為病中的家人代禱' }
    ]
  );
});

test('proofreading response is limited to requested fields and never mutates source text', () => {
  const requested = [{ id: 'announcements.0', text: '請按時到教會' }];
  const result = proofreader.normalizeProofreadingResponse({
    suggestions: [
      { id: 'announcements.0', suggestion: '請按時到教會。', changed: true },
      { id: 'unknown', suggestion: '不應顯示' }
    ]
  }, requested);

  assert.deepEqual(result, [{
    id: 'announcements.0',
    text: '請按時到教會',
    suggestion: '請按時到教會。',
    changed: true,
    note: ''
  }]);
  assert.equal(requested[0].text, '請按時到教會');
});

test('report page includes the AI proofreading entry point and shared ministry API config', () => {
  const reportPath = path.join(__dirname, '..', 'apps', 'LKC_SundayBulletin', 'reports.html');
  const html = fs.readFileSync(reportPath, 'utf8');
  assert.match(html, /id=["']btnAiProofread["']/);
  assert.match(html, /reports-ai\.js/);
  assert.match(html, /_GAS_KEY\s*=\s*["']LKC_MinistrySchedule["']/);
});

test('GAS route exposes a proofreading action backed by MinistryCore', () => {
  const corePath = path.join(__dirname, '..', 'scratch_gas_sunday', 'Core.js');
  const ministryPath = path.join(__dirname, '..', 'scratch_gas_sunday', 'MinistryCore.js');
  const core = fs.readFileSync(corePath, 'utf8');
  const ministry = fs.readFileSync(ministryPath, 'utf8');
  assert.match(core, /ministry_proofreadFields/);
  assert.match(ministry, /function ministry_proofreadFields\s*\(/);
  assert.match(ministry, /callGemini\(systemPrompt, userText, \{ useCache: true \}\)/);
});
