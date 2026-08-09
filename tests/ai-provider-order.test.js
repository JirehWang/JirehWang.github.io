const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadProviderOrder(rows) {
  const sourcePath = path.join(__dirname, '..', 'scratch_gas_sunday', 'GeminiHelper.js');
  const source = fs.readFileSync(sourcePath, 'utf8') +
    '\nthis.__test = { getProviderOrder: _getAiProviderOrder };';
  const context = {
    console,
    SpreadsheetApp: {
      openById() {
        return {
          getSheetByName() {
            return { getDataRange: () => ({ getValues: () => rows }) };
          }
        };
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context.__test.getProviderOrder;
}

test('AI provider fallback order follows API key rows in AI_Config', () => {
  const getProviderOrder = loadProviderOrder([
    ['OPENROUTER_API_KEY', 'openrouter-key'],
    ['OPENROUTER_MODEL', 'deepseek/deepseek-chat'],
    ['GEMINI_API_KEY', 'gemini-key'],
    ['GEMINI_MODEL', 'gemini-3.1-flash-lite'],
    ['NVIDIA_API_KEY', 'nvidia-key'],
    ['NVIDIA_MODEL', 'google/gemma-3n-e2b-it']
  ]);

  assert.deepEqual(Array.from(getProviderOrder()), ['OpenRouter', 'Gemini', 'NVIDIA']);
});
