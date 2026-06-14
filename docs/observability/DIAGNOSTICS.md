# LKC Diagnostics Plan

This is the next backend step after the existing Firebase log dashboard.

## Minimum Health Action

Each GAS project should expose a read-only diagnostic action:

```text
health_check
```

Return:

```json
{
  "status": "success",
  "system": "LKC_Group",
  "time": "2026-06-14T00:00:00.000Z",
  "checks": {
    "sheetOpen": true,
    "cacheService": true,
    "firebaseAuth": true,
    "scriptProperties": true
  },
  "durationMs": 123
}
```

## Suggested Checks

- `sheetOpen`: can open the configured spreadsheet and read one cell.
- `cacheService`: can write/read/remove a short diagnostic key.
- `firebaseAuth`: can get a token or read/write a diagnostic cache path when the project owns Firebase sync.
- `scriptProperties`: required secrets/config keys exist, without returning their values.
- `actionRouter`: required actions are registered.

## Stability Bot Integration

1. `logs.html` remains the human dashboard and export surface.
2. Export filtered log JSON from `logs.html`.
3. `church_logs_summary.py --input exported.json` produces the log evidence summary.
4. `health_check` results become `metrics` and `chaos_tests` evidence for `arch_stability_bot.py`.
5. The stability bot produces the current score and next actions.

## Do Not Log

- Tokens.
- Service account JSON.
- Phone numbers.
- Full member data.
- Full request payloads.

Keep log metadata short and structural.
