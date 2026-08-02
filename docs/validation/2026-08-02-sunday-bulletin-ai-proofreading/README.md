# Sunday bulletin AI proofreading validation

## Scope

- Add a proofreading suggestion beside every editable report field.
- Keep the original report field unchanged; suggestions are display-only.
- Reuse the ministry system's `ministry_*` GAS route and shared `GeminiHelper` provider fallback.

## Acceptance criteria

1. Blank fields are not sent to the AI service.
2. Each returned suggestion is matched to a requested field id only.
3. The original value is never overwritten by the AI response.
4. AI failure leaves the original form usable and reports an error state.
5. The browser sends no provider API key; the GAS backend owns model access.

## Validation stages

| Stage | Status | Evidence |
| --- | --- | --- |
| SDD | required | This scope and acceptance list |
| DDD | required | Existing MinistryCore/GeminiHelper service boundary reused |
| BDD | required | UI contract: check fields, show suggestions, preserve originals |
| TDD | required | `tests/sunday-bulletin-ai-proofreading.test.js` red-green cycle |
| Integration | required | Frontend request contract plus GAS route/source checks |
| E2E | exempted | Live AI calls require deployed GAS credentials; use browser/manual checks for UI |
| Completion verification | required | Fresh test and syntax output before handoff |
