# Project Journal

## Project Snapshot

- Project: LKC1958_June_1.github.io
- Root: `D:/py/LKC1958_June_1.github.io-worship-cloud`
- Contract: `project_contract.yml`
- Current focus: The Taiwanese worship PPT frontend is Firebase-first for synchronized content and retains GAS fallback until the GAS project begins populating the new RTDB/Storage paths.
- Last updated: 2026-07-16

## Stable Facts

- App entrypoint: `apps/LKC_WorshipPPT/index.html`.
- Existing browser draft key: `lkc-taiwanese-worship-draft`.
- Layout data shape: `layoutState.groups`, `layoutState.pageAssignments`, `layoutState.hymnOpacityBySection`, and `layoutState.outputScale`.
- Firebase project initialization already exists in `firebase/firebase-config.js` and includes RTDB.
- Verification: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1`.

## Open Risks

- Firebase Email/Password authentication and the dedicated editor account must be configured in Firebase Console before protected cloud writes can succeed.
- Database rules must be merged with the deployed rules rather than overwriting unrelated production paths blindly.
- In-app browser automation could not initialize in this environment (`Cannot redefine property: process`), so desktop/mobile visual acceptance remains manual.

## Recent Entries

### 2026-07-16

- Changed: Renamed the application from `LKC_TaiwaneseWorshipPPT` to `LKC_WorshipPPT` and the user-facing product to `禮拜PPT產生器`, updating the admin entry, workflow contract, verification path, architecture, and Firebase documentation. The longer brand is kept on one line at desktop and narrow widths. The current Taiwanese worship flow and its existing Firebase/local draft keys remain unchanged so weekly content and shared settings keep working.
- Changed: Separated page browsing from layout-page selection in the worship flow navigator. Clicking a page row now previews it without changing its checkbox; checking a page still selects it for layout editing and jumps to that page, while checking a chapter retains bulk selection and first-page preview. Chapter names remain native expand/collapse controls, with distinct current-page and selected-page styling plus a short usage hint.
- Changed: Updated the sermon title page to render `講道：{講道題目}` as the heading, with only the speaker and scripture in the body box. The shared production geometry now measures wrapped heading lines and vertically centers the complete heading/body group identically in canvas preview and PowerPoint export.
- Focus: Final alignment of the praise and sermon title pages between the browser canvas and exported PowerPoint.
- Changed: Made the sermon a dedicated title page and grouped each page's secondary information into the same single `.body` box used by the existing layout measurement and cloud-parameter flow. Praise now keeps song title and performer together; sermon keeps topic, speaker, and scripture together.
- Changed: Extended the existing centered section calculation by detail-line count. The title/body group stays vertically centered at 50% while two-line praise and three-line sermon content receive proportional body heights in both canvas and export.
- Changed: Resolved the subsequent `main` merge conflicts by moving that line-count calculation into `slide-production.js`'s shared `resolvedLayoutForPage()` path. This preserves `main`'s deterministic wrapping, title anchoring, hymn-overlay, and preview-entry fixes while keeping praise and sermon centered in both preview and export.
- Verification: Browser visual QA measured a 50.0% group center for both pages and confirmed the sermon speaker remains inside the canvas. The focused export regression test verifies the exported text groups use the same center and increasing body height for additional lines.

### 2026-07-15

- Focus: Shared cloud layout and password-gated editing.
- Daily summary — Firebase and access control: Moved the church-wide layout, score opacity, and text/image output-scale settings into Firebase RTDB. Editing and saving are locked behind Firebase Email/Password authentication using the dedicated worship layout account; localStorage remains only for first migration and offline fallback.
- Daily summary — Export behavior: Preserved imported hymn and responsive-reading pages as centered raster images, kept reports and other generated pages editable, restored source slide dimensions, and added independent 80-120% text/image output scaling in the top toolbar.
- Daily summary — Layout corrections: Fixed invalid default bounds, preserved imported coordinates, aligned title anchors and centered templates with the browser preview, restored the first responsive-reading title size, included hymn/doxology song names, and centered praise lyrics in the full safe area.
- Daily summary — Data migration preparation: Added Firebase-first reads for calendar, Bible, PPT library, reports, and praise content, with GAS fallback while Firebase records are absent. Added RTDB/Storage rule templates and the Firebase content synchronization contract; no GAS code was changed.
- Daily summary — Quality and delivery: Expanded automated coverage to 50 tests, added JavaScript/rules validation, and visually verified representative real PPTX exports for score images, editable reports, centered titles, hymn names, and praise lyrics.
- Changed: Added cloud-authoritative RTDB layout storage, Firebase Auth password dialog, locked controls, first-use local migration, offline backup behavior, rules template, setup guide, workflow contract, and verification script.
- Changed: Fixed PPTX export for ungrouped pages by merging safe default bounds before coordinate conversion; added a regression test proving every exported text box has finite, positive dimensions.
- Changed: Kept imported PowerPoint text in its source coordinates unless the page has an explicit stored layout group; generated pages continue to receive safe default bounds.
- Changed: Forwarded the live model, background color, and background image from `app.js` into the standalone exporter, and bumped browser cache versions for both changed scripts.
- Changed: Rasterized imported hymn and responsive-reading pages to one transparent PNG per page while leaving reports, liturgy, scripture, and other generated pages as editable PowerPoint text.
- Changed: Restored the original `13.333 × 7.5` wide-slide canvas so font point sizes and percentage positions use the same physical dimensions as the source library.
- Changed: Stored each score-related section's white-overlay opacity in the shared Firebase layout document. The sliders and sync control now use the same Firebase Auth lock as layout editing.
- Changed: Added password-protected, church-wide output percentages for editable text and rasterized images. Text scaling changes font sizes without moving text boxes; image scaling keeps hymn and responsive-reading pages centered. Both settings accept 80-120% and default to 100%.
- Changed: Fixed the first responsive-reading page so placeholder text inherits its font size from the linked PowerPoint slide layout. This restores the source title from the 18pt parser fallback to its intended 60pt size without changing explicitly sized text.
- Changed: Moved the shared text/image output percentages out of the floating layout-parameter editor and into the top toolbar. The two fields and their save button remain protected by the same Firebase Auth lock and continue storing one church-wide setting.
- Changed: Matched explicitly positioned PowerPoint title boxes to the browser preview's top anchoring. Tall shared title boxes no longer vertically center their text into the body area, fixing the creed pages reproduced in the 2026-07-15 export.
- Changed: Added template-specific PowerPoint defaults for ungrouped cover and section pages. Their title/subtitle groups now remain vertically centered like the browser preview, and section subtitles use the preview's 36pt base size instead of the generic 48pt body size.
- Changed: Added a Firebase-first content reader for service calendar data, Bible query results, PPT library index, reports, and praise records. Missing or failed Firebase reads preserve the existing GAS/JSONP path. PPT library entries with a Firebase Storage `downloadUrl` or `storageUrl` bypass the GAS Base64 endpoint.
- Changed: Added RTDB and Storage rule templates plus a synchronization contract for `worshipPpt/content` and `worshipPpt/library`. The RTDB layout rule template now also validates protected opacity and text/image output-scale fields.
- Changed: Completed hymn-family section titles by reading the PPT Library song name from the loaded section model when the generated section entry has no kicker. This restores names on hymn 1, hymn 2, pre-service hymns, and doxology title pages.
- Changed: Matched ungrouped praise lyric pages to the browser template's 10% full-slide safe area and vertical centering instead of centering inside the lower generic body box.
- Learned: The full browser draft combines layout state with content and background image data; only `layoutState.groups` and `layoutState.pageAssignments` belong in the shared RTDB node.
- Learned: Browser preview CSS supplied implicit defaults, while `ppt-export.js` previously converted missing layout values into `NaN`; PptxGenJS serialized those text boxes with zero-size OOXML bounds.
- Learned: Applying those same generated-page defaults to an ungrouped `ppt-import` page resized all imported text into the default content rectangle, causing the hymn and score pages to overlap. The export click handler also omitted lexical background state because top-level `let` bindings are not properties of `window`.
- Learned: The generic generated-page fallback also placed ungrouped `cover` and `section` titles at `Y=6%`, while their browser templates use flex centering. These page kinds require their own fallback geometry whenever no cloud layout bounds are assigned.
- Learned: The live RTDB contained layout groups and page assignments, but not `hymnOpacityBySection` or `outputScale`. It had only three short-lived `cal_getEvents` cache keys for other dates; 2026-07-15, PPT library index/files, Bible queries, reports, and praise content were absent. `cal_getPptLibraryIndex`, `cal_getPptLibraryFile`, and `cal_queryBible` were not configured as Firebase cacheable actions in `config.js`.
- Learned: PPT Library integration writes parsed song names to `model[sectionId].kicker`, while section export previously inspected only the flattened deck entry. Praise lyrics already used middle anchoring, but their generic `Y=24%, H=68%` box had a center at 58% of the slide rather than the browser template's 50%.
- Verification: `scripts/verify.ps1` passed all 50 app tests, JavaScript syntax checks, and rules JSON validation. Firebase-first tests cover stable content paths, hit/miss behavior, GAS bypass on a Firebase hit, and Storage URL priority. A real PptxGenJS export visually confirmed hymn 65 and doxology 510 title pages include their song names and a four-line praise lyric block is centered in the full safe area. A prior real export at 90%/90% produced a centered rasterized hymn page with even margins and a separate editable report whose font sizes were reduced without moving its text boxes. The live responsive-reading 21 source deck was also parsed and rendered before/after the placeholder fix, confirming its first-page title changed from the incorrect 18pt fallback to the inherited 60pt source size. The title-anchor regression test uses the exact tall title/body bounds found on slides 13-15 of the 2026-07-15 attachment. A separate real export of cover, title-only, and title-plus-subtitle pages was rendered after the centered-template fix and visually confirmed all three groups at the slide center.
- Next: Enable Email/Password Auth, create the dedicated editor account with the separately agreed password, merge/deploy RTDB rules, then perform desktop/mobile acceptance and first migration.
