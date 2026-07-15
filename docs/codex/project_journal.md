# Project Journal

## Project Snapshot

- Project: LKC1958_June_1.github.io
- Root: `D:/py/LKC1958_June_1.github.io-worship-cloud`
- Contract: `project_contract.yml`
- Current focus: Shared Firebase layout is active; PPTX export fidelity fixes are verified and awaiting deployment acceptance.
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
- Changed: Fixed PPTX export for ungrouped pages by merging safe default bounds before coordinate conversion; added a regression test proving every exported text box has finite, positive dimensions.
- Changed: Kept imported PowerPoint text in its source coordinates unless the page has an explicit stored layout group; generated pages continue to receive safe default bounds.
- Changed: Forwarded the live model, background color, and background image from `app.js` into the standalone exporter, and bumped browser cache versions for both changed scripts.
- Learned: The full browser draft combines layout state with content and background image data; only `layoutState.groups` and `layoutState.pageAssignments` belong in the shared RTDB node.
- Learned: Browser preview CSS supplied implicit defaults, while `ppt-export.js` previously converted missing layout values into `NaN`; PptxGenJS serialized those text boxes with zero-size OOXML bounds.
- Learned: Applying those same generated-page defaults to an ungrouped `ppt-import` page resized all imported text into the default content rectangle, causing the hymn and score pages to overlap. The export click handler also omitted lexical background state because top-level `let` bindings are not properties of `window`.
- Verification: `scripts/verify.ps1` passed all 40 app tests, JavaScript syntax checks, and rules JSON validation. A real PptxGenJS export rendered imported text at its original bounds over the requested non-white background.
- Next: Enable Email/Password Auth, create the dedicated editor account with the separately agreed password, merge/deploy RTDB rules, then perform desktop/mobile acceptance and first migration.
