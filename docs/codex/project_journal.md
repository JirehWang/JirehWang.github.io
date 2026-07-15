# Project Journal

## Project Snapshot

- Project: LKC1958_June_1.github.io
- Root: `D:/py/LKC1958_June_1.github.io-worship-cloud`
- Contract: `project_contract.yml`
- Current focus: Imported score/response pages are image-based; editable generated pages, protected shared opacity, and separate output scaling are verified and awaiting deployment acceptance.
- Last updated: 2026-07-15

## Stable Facts

- App entrypoint: `apps/LKC_TaiwaneseWorshipPPT/index.html`.
- Existing browser draft key: `lkc-taiwanese-worship-draft`.
- Layout data shape: `layoutState.groups`, `layoutState.pageAssignments`, `layoutState.hymnOpacityBySection`, and `layoutState.outputScale`.
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
- Changed: Rasterized imported hymn and responsive-reading pages to one transparent PNG per page while leaving reports, liturgy, scripture, and other generated pages as editable PowerPoint text.
- Changed: Restored the original `13.333 × 7.5` wide-slide canvas so font point sizes and percentage positions use the same physical dimensions as the source library.
- Changed: Stored each score-related section's white-overlay opacity in the shared Firebase layout document. The sliders and sync control now use the same Firebase Auth lock as layout editing.
- Changed: Added password-protected, church-wide output percentages for editable text and rasterized images. Text scaling changes font sizes without moving text boxes; image scaling keeps hymn and responsive-reading pages centered. Both settings accept 80-120% and default to 100%.
- Learned: The full browser draft combines layout state with content and background image data; only `layoutState.groups` and `layoutState.pageAssignments` belong in the shared RTDB node.
- Learned: Browser preview CSS supplied implicit defaults, while `ppt-export.js` previously converted missing layout values into `NaN`; PptxGenJS serialized those text boxes with zero-size OOXML bounds.
- Learned: Applying those same generated-page defaults to an ungrouped `ppt-import` page resized all imported text into the default content rectangle, causing the hymn and score pages to overlap. The export click handler also omitted lexical background state because top-level `let` bindings are not properties of `window`.
- Verification: `scripts/verify.ps1` passed all 42 app tests, JavaScript syntax checks, and rules JSON validation. A real PptxGenJS export at 90%/90% produced a centered rasterized hymn page with even margins and a separate editable report whose font sizes were reduced without moving its text boxes.
- Next: Enable Email/Password Auth, create the dedicated editor account with the separately agreed password, merge/deploy RTDB rules, then perform desktop/mobile acceptance and first migration.
