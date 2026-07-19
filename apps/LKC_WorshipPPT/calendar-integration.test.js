const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, 'calendar-integration.js'), 'utf8');

test('shows one non-blocking warning dialog that names missing import sources', () => {
  assert.match(source, /buildMissingSourceReminders\(\{ date, event, model, bulletinResult, libraryResults, profile \}\)/);
  assert.match(source, /reminders\.length[\s\S]{0,120}window\.alert\(/);
  assert.match(source, /formatMissingSourceReminder\(reminders\)/);
  assert.match(source, /status\(`\$\{calendarSummary\}/);
});
