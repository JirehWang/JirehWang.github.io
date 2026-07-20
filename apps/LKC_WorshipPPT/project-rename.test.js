const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..', '..');
const newApp = path.join(root, 'apps', 'LKC_WorshipPPT');
const oldApp = path.join(root, 'apps', 'LKC_TaiwaneseWorshipPPT');

test('uses only the generic LKC_WorshipPPT application directory', () => {
  assert.equal(fs.existsSync(oldApp), false);
  assert.equal(fs.existsSync(path.join(newApp, 'index.html')), true);
});

test('uses 禮拜PPT產生器 as the user-facing product name', () => {
  const index = fs.readFileSync(path.join(newApp, 'index.html'), 'utf8');
  const admin = fs.readFileSync(path.join(root, 'admin.html'), 'utf8');
  assert.match(index, /<title>禮拜PPT產生器<\/title>/);
  assert.match(index, /林口教會 <span>\/ 禮拜PPT產生器<\/span>/);
  assert.match(admin, /href="apps\/LKC_WorshipPPT\/"[\s\S]*?<div class="title">禮拜PPT產生器<\/div>/);
});

test('project workflow and architecture references use the renamed path', () => {
  const contract = fs.readFileSync(path.join(root, 'project_contract.yml'), 'utf8');
  const verify = fs.readFileSync(path.join(root, 'scripts', 'verify.ps1'), 'utf8');
  const relation = fs.readFileSync(path.join(root, 'docs', 'SYSTEM_RELATION_GRAPH.md'), 'utf8');
  const architecture = fs.readFileSync(path.join(newApp, 'ARCHITECTURE.md'), 'utf8');
  for (const content of [contract, verify, relation]) {
    assert.doesNotMatch(content, /LKC_TaiwaneseWorshipPPT/);
  }
  assert.match(contract, /apps\/LKC_WorshipPPT\/index\.html/);
  assert.match(verify, /apps\\LKC_WorshipPPT/);
  assert.match(verify, /All Worship PPT generator tests passed\./);
  assert.doesNotMatch(verify, /Taiwanese Worship PPT/);
  assert.match(architecture, /^# 禮拜PPT產生器/m);
});

test('architecture documents the shared core and active template boundaries', () => {
  const architecture = fs.readFileSync(path.join(newApp, 'ARCHITECTURE.md'), 'utf8');
  const relation = fs.readFileSync(path.join(root, 'docs', 'SYSTEM_RELATION_GRAPH.md'), 'utf8');

  for (const templateId of ['taiwanese', 'joint-taiwanese', 'joint-mandarin', 'mandarin']) {
    assert.match(architecture, new RegExp('`' + templateId + '`'));
  }

  assert.match(architecture, /worshipPpt\/layoutConfig\/templates\/\{templateId\}/);
  assert.match(architecture, /worshipPpt\/content\/services\/\{date\}\/\{templateId\}/);
  assert.match(architecture, /共用的核心/);
  assert.match(architecture, /現行 template profile/);
  assert.match(relation, /ARCHITECTURE\.md/);
  assert.match(relation, /多模板擴充邊界（台語、聯合台語與聯合華語已實作）/);
});
