# External GAS v2 cache synchronization

This change is a deployment packet for three independent Apps Script projects.
It is deliberately **not** applied by this repository: the source projects live
under `D:\program\LKC` and require their own local write permission and Apps
Script authorization.

## Bounded source map

| Project | clasp script ID | read topics | writer ownership | v2 cache owner | reconciliation installer |
| --- | --- | --- | --- | --- | --- |
| 主日出席_測試版 | `1A_dni5PbG5CMrIzlSWfk0PLhzkF4PX65ecRsaruYb-pW4kq0VJYFsEJl` | Sunday/group/ministry/calendar/worship routes; `memberStatus_*` remains bypassed | `Core.js`, with committed writers in `GroupCore.js` and the existing domain modules | `FirebaseSync.js`: `firebaseCacheWriteThrough`, `firebaseInvalidate`, `firebaseReconcilePendingTopics` | `GroupCore.js: setupKeepWarmTrigger()` (daily) |
| 兒童出席_GAS | `1zAXSMNBc1L9_H_UGTFC6sYuEYA9jjNEloQyks7l-eyEn9vyZgq2FLTcS` | `children_getAllMembers`, `children_getSmartAttendanceList`, `children_getGroupConfig`, statistics/trend/chart | `Core.js` dispatches to `MemberDB.js` / `AttendanceDB.js`; those modules invalidate only after their Sheet commits | `FirebaseSync.js`: prefixed v2 entry writer and pending repair | `FirebaseSync.js: setupKeepWarmTrigger()` (daily) |
| 新家人管理系統 | `1kJmGfYaliCggqJfTf_k4GMlpUPJ81mpZBla2jDskE84SPlMj57TsTYLl` | `getTrackingCases`, `getClosedCases` | `Code.js` write functions commit under `LockService`, then call `refreshNewFamilyCaches_` | `Code.js`: server-owned list rebuild with v2 metadata and pending repair | `KeepWarm.js: setupKeepWarmTrigger()` (daily) |

`memberStatus_*` is intentionally not added to the shared-cache allowlist. Its
five-minute aggregate is derived from several Sheets and no safe targeted
write-through exists in the scoped source. It must continue to use the
front-end direct-GAS bypass until a separate aggregate-rebuild contract is
implemented and tested.

## Cache invariants in the patch

1. Browser callers never write or delete shared Firebase entries.
2. A GAS read captures the topic generation before reading Sheets and publishes
   only if that generation is still current.
3. A successful Sheet mutation invalidates/bump-generates its affected topics.
4. Firebase failures create a bounded pending-repair marker and do not turn a
   successful mutation into a second Sheet read or a second API call.
5. All v2 entries contain `schemaVersion: 2`, `generation`, `sourceRevision`,
   `updatedAt`, and `expiresAt: null` (long-lived data, not traffic TTL).
6. The installers are idempotent but are not run by this patch.

## Apply and verify after local permission is granted

From `D:\program\LKC`, first verify the patch against the exact source mapped
above:

```powershell
git apply --check D:\program\Github\LKC1958_June_1.github.io\docs\validation\2026-08-11-external-gas-cache-v2\external-gas-v2-sync.patch
git apply D:\program\Github\LKC1958_June_1.github.io\docs\validation\2026-08-11-external-gas-cache-v2\external-gas-v2-sync.patch
node --check 主日出席_測試版\Core.js
node --check 主日出席_測試版\FirebaseSync.js
node --check 主日出席_測試版\GroupCore.js
node --check 兒童出席_GAS\Core.js
node --check 兒童出席_GAS\FirebaseSync.js
node --check 新家人管理系統\Code.js
node --check 新家人管理系統\KeepWarm.js
```

Then, only after source review and clasp authentication, inspect each project
with `npx @google/clasp status` from its own directory. Do not run `clasp push`
or an Apps Script deployment until the release owner approves it.

After a permitted deploy, run the existing installer **once per project** from
the Apps Script editor and approve its Script/Spreadsheet/UrlFetch scopes:

```text
主日出席_測試版: setupKeepWarmTrigger()
兒童出席_GAS: setupKeepWarmTrigger()
新家人管理系統: setupKeepWarmTrigger()
```

Production E2E is still mandatory: prove a Firebase hit, a single GAS fallback,
a successful mutation followed by an updated cache entry, a stale-generation
write rejection, and a Firebase failure that is repaired by the next daily
reconciliation. The patch alone does not grant these authorizations or prove
them against real Firebase/Sheets.
