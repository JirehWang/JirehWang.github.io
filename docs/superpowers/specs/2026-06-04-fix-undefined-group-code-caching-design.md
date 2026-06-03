# Design Specification: Fix Undefined Group Code URL Parameter by Removing Caching

## 1. Problem Description
After the introduction of backend-encrypted group URLs in commit `59d9a1e`, the frontend redirects users using the `encryptedCode` property returned by the backend. However, because `findGroupByCode` and `verifyGroup` are cacheable in `config.js`, stale cached responses from Firebase lacking the `encryptedCode` property are returned, leading to `code=undefined` in URL redirects. Furthermore, caching authentication/login actions introduces security issues and synchronization delay when group passwords/configurations change.

## 2. Proposed Changes

### 2.1. [config.js](file:///d:/program/Github/LKC1958_June_1.github.io/config.js)
We will remove `findGroupByCode` and `verifyGroup` from the `_CACHEABLE_ACTIONS` list. This ensures that:
- Verification and group lookup requests bypass the Firebase cache entirely and query the Google Apps Script directly.
- The `encryptedCode` is always fetched live, resolving the `undefined` parameter bug.
- Password/configuration updates are immediately effective.

### 2.2. Firebase Cache Purge
We will execute a one-time cleanup script to delete the keys under `cache/findGroupByCode` and `cache/verifyGroup` in the Firebase Realtime Database to purge the stale cache entries.

## 3. Verification Plan
- **Verification via Live Queries**: We will run tests using our python script to verify that requests to `findGroupByCode` are fetched dynamically and no new cache entries are created for these endpoints.
- **Manual Verification**: Verify that log-in redirects to `group.html` now correctly include the `enc_` prefixed code parameter.
