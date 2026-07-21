const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const appSource = fs.readFileSync(path.join(__dirname, 'app.js'), 'utf8');
const indexSource = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');

assert.match(appSource, /let selectedImageFile = null;/, 'image selection state must be declared');
assert.match(appSource, /dropzone\.ondrop\s*=\s*\(e\)\s*=>/, 'drag-and-drop handler must be registered');
assert.match(appSource, /fileInput\.onchange\s*=\s*\(e\)\s*=>/, 'file picker handler must be registered');
assert.match(appSource, /churchAPI\('cal_parsePrayerImage'/, 'image parsing must go through the GAS proxy');
assert.doesNotMatch(appSource, /generativelanguage\.googleapis\.com/, 'browser code must not call Gemini directly');
assert.match(indexSource, /\.dialog-tab-content\s*\{\s*display:\s*none;/,
  'inactive import tabs must be hidden');
assert.match(indexSource, /\.dialog-tab-content\.active\s*\{\s*display:\s*block;/,
  'active import tab must be visible');
assert.match(indexSource, /@media \(max-width:\s*620px\)/,
  'PrayerPPT must define a mobile layout');
assert.match(indexSource, /window\._GAS_KEY\s*=\s*['"]LKC_PrayerPPT['"]/,
  'PrayerPPT must use its dedicated main-GAS route');

console.log('PrayerPPT AI integration source checks passed');
