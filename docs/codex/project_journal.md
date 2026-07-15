# Project Journal

## Project Snapshot

- Project: LKC1958_June_1.github.io
- Root: `D:/py/LKC1958_June_1.github.io-worship-cloud`
- Contract: `project_contract.yml`
- Current focus: Shared Firebase layout implementation complete; Firebase Console account/rules activation and browser acceptance remain.
- Last updated: 2026-07-15

## Stable Facts

- App entrypoint: `apps/LKC_TaiwaneseWorshipPPT/index.html`.
- Existing browser draft key: `lkc-taiwanese-worship-draft`.
- Layout data shape: `layoutState.groups` and `layoutState.pageAssignments`.
- Firebase project initialization already exists in `firebase/firebase-config.js` and includes RTDB.
- Verification: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1`.

## Open Risks

- Firebase Email/Password authentication and the dedicated editor account must be configured in Firebase Console before protected cloud writes can succeed.
- Database rules must be merged with the deployed rules rather than overwriting unrelated production paths blindly.
- In-app browser automation could not initialize in this environment (`Cannot redefine property: process`), so desktop/mobile visual acceptance remains manual.

## Recent Entries

### 2026-07-15

- Focus: Shared cloud layout and password-gated editing.
- Changed: Added cloud-authoritative RTDB layout storage, Firebase Auth password dialog, locked controls, first-use local migration, offline backup behavior, rules template, setup guide, workflow contract, and verification script.
- Learned: The full browser draft combines layout state with content and background image data; only `layoutState.groups` and `layoutState.pageAssignments` belong in the shared RTDB node.
- Verification: `scripts/verify.ps1` passed all 37 app tests, JavaScript syntax checks, and rules JSON validation. Local HTTP route returned 200 with the expected script order. Browser visual automation was unavailable due a runtime initialization error.
- Next: Enable Email/Password Auth, create the dedicated editor account with the separately agreed password, merge/deploy RTDB rules, then perform desktop/mobile acceptance and first migration.
