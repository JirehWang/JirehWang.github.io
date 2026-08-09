# Attendance Firebase Temp Validation

Date: 2026-08-04 (Asia/Taipei)

## Contract

- Browser writes are acknowledged by Firebase before a scan is treated as accepted.
- A retry keeps the same `scope + UID` child and `requestId`; it cannot append a duplicate staging row.
- `checked=false` is a short-lived tombstone so an uncheck survives until the next batch.
- GAS batches one scope every 5 seconds, applies a `SYNC_TEMP` mutation plan under `LockService`, then uses Firebase ETag/`If-Match` for conditional acknowledgement.
- The local retry queue is bounded at 100 items; the QR writer sends at most 8 Firebase writes per attempt.

## Local evidence

- Frontend contract tests: 4 passed.
- GAS mutation-plan tests: 3 passed.
- Full frontend suite: 55 passed, 1 skipped, 0 failed.
- Full GAS suite: 27 passed, 0 failed.
- JavaScript syntax checks passed for the changed attendance, AGM, scanner, GAS attendance, GAS temp, and Core files.

## Deployment prerequisites

This change has not been deployed to the formal GAS Web App or Firebase Rules in this turn. Before production use:

1. Enable Firebase Anonymous Authentication.
2. Deploy `firebase/database.rules.attendance-temp.json` to the RTDB ruleset.
3. Deploy the new `AttendanceTemp.js` and `Core.js` route to GAS.
4. Publish the updated GitHub Pages assets.
5. Re-run the controlled checks: three devices in one scope, two scopes with two devices each, 100 scans at 3-second intervals, and verify zero missing rows after final flush.

