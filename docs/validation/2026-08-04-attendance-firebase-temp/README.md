# Attendance Firebase Temp Validation

Date: 2026-08-04 (Asia/Taipei)

## Contract

- Browser writes are acknowledged by Firebase before a scan is treated as accepted.
- A retry keeps the same `scope + UID` child and `requestId`; it cannot append a duplicate staging row.
- `checked=false` is a 6-hour tombstone so an uncheck survives until the next batch and cannot be resurrected by a stale writer.
- Every pending UID carries a monotonic `revision`, `ownerId`, `source`, and a 10-minute `lockedUntil` lease. The owner can cancel immediately; another device cannot cancel while the lease is active.
- GAS batches one scope every 30 seconds, reads Firebase after acquiring `LockService`, applies a revision-guarded `SYNC_TEMP` mutation plan, then uses Firebase ETag/`If-Match` for conditional acknowledgement.
- The local retry queue is bounded at 100 items; the QR writer sends at most 8 Firebase writes per attempt.

## Local evidence

- Attendance state and contract tests: 13 passed.
- GAS mutation-plan tests: 3 passed.
- Full repository suite: 85 passed, 1 skipped, 0 failed (86 tests).
- JavaScript syntax checks passed for the changed attendance, AGM, scanner, GAS attendance, GAS temp, and Core files.

## Deployment prerequisites

This change has not been deployed to the formal GAS Web App or Firebase Rules in this turn. Before production use:

1. Enable Firebase Anonymous Authentication.
2. Deploy `firebase/database.rules.attendance-temp.json` to the RTDB ruleset.
3. Deploy the new `AttendanceTemp.js` and `Core.js` route to GAS.
4. Publish the updated GitHub Pages assets.
5. Re-run the controlled checks: three devices in one scope, two scopes with two devices each, 100 scans at 3-second intervals, and verify zero missing rows after final flush.
