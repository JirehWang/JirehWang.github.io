# LKC Observability Runbook

Use this when `logs.html` shows critical/error events, slow API calls, stale cache symptoms, or users report data not updating.

## First Check

1. Open `logs.html`.
2. Select today's date.
3. Filter by `critical` and `error`.
4. Note the affected `system`, `action`, `message`, `durationMs`, and `meta`.
5. Click the export button in `logs.html` to download the current filtered log JSON.
6. Run the local summary against that exported file:

```powershell
C:\Users\jireh\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe D:\program\py\arch_stability_bot\church_logs_summary.py --input D:\Downloads\lkc-logs-YYYY-MM-DD.json
```

Paste the resulting JSON into `church_system_architecture.json` under `log_summary`, then run:

```powershell
C:\Users\jireh\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe D:\program\py\arch_stability_bot\arch_stability_bot.py D:\program\py\arch_stability_bot\church_system_architecture.json --pretty
```

## Critical Error Triage

### API request failed

Likely causes:

- GAS deployment unavailable or slow.
- Network request blocked or timed out.
- Token/action mismatch.
- Downstream Sheets/Firebase/UrlFetch failure.

Actions:

1. Retry the same action from the app once.
2. Check whether the same `system/action` repeats in logs.
3. If repeated, bypass cache by using the direct GAS action path or temporarily remove the action from cacheable actions.
4. If only one app fails, inspect that app's `_GAS_KEY`, action prefix, and API URL mapping.

### Invalid cache response / fallback to GAS

Likely causes:

- Stale Firebase cache schema.
- Old cache missing a new response field.
- Login/verification action was cached when it should be live.

Actions:

1. Purge only the affected cache topic when possible.
2. Verify the action is still safe to cache.
3. Confirm the fallback GAS call rebuilds a valid cache response.
4. If authentication or verification is involved, keep it uncached.

### Slow API

Signals:

- `durationMs > 3000` for direct GAS calls.
- `durationMs > 5000` or repeated warnings.
- p95 duration elevated in `church_logs_summary.py`.

Actions:

1. Check whether the action was cache hit, cache miss, or direct GAS.
2. Inspect Sheet access paths for repeated `openById`, full-sheet reads, or cross-spreadsheet calls.
3. Add or tune CacheService/Firebase cache only for read-only actions.
4. Add a smaller response shape for dashboards or autocomplete endpoints.

## Degraded Mode

If Firebase RTDB has issues:

- Read-only cache may fail; apps should fall back to GAS.
- Expect slower responses and higher GAS quota usage.
- Avoid manual full-cache purge during peak use unless stale data is worse than slowness.

If GAS has issues:

- Cached read-only pages may continue working until TTL expires.
- Writes may fail or be delayed.
- Use manual Google Sheet operation for urgent attendance/member updates.

If Google Sheets has issues:

- Treat this as the main data-store outage.
- Avoid retries that create duplicate writes.
- Prefer read-only communication to users until Sheet access recovers.

## Recovery Confirmation

The incident is stable when:

- `critical_count = 0` for the latest window.
- Repeated `error_count` stops increasing for the affected action.
- p95 duration returns below 3000ms for normal interactions.
- Cache fallback events stop repeating after a valid cache rebuild.
- A user can complete the failed workflow end to end.
