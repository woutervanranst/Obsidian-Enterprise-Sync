# Architecture & Technical Details

This document describes how **OneDrive + Enterprise SharePoint + Teams Sync** (`entra-sharepoint-sync`) synchronizes Obsidian vaults directly against Microsoft Graph without intermediate servers.

---

## 1. Unified Microsoft Graph Drive Model

Consumer OneDrive, OneDrive for Business, SharePoint Online Document Libraries, and Microsoft Teams channel files all expose the same underlying resource in Microsoft Graph: a **Drive** (`/v1.0/drives/{driveId}`) and a **Root Folder Item** (`/items/{rootItemId}`).

Once a vault target is resolved to `{ driveId, rootItemId }`, all synchronization, conflict detection, version history, and sharing operations use identical endpoints:

| Capability | Consumer OneDrive (`common`) | Enterprise SharePoint & Teams (`{tenantId}`) |
|---|---|---|
| **OAuth2 Authority** | `https://login.microsoftonline.com/common/oauth2/v2.0` | `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0` |
| **Vault Discovery** | `/me/drive/root:/{folder}:/children` & `/me/drive/sharedWithMe` | `/me/followedSites` → `/sites/{siteId}/drive` & `/me/drive/sharedWithMe` |
| **Delta Sync** | `GET /drives/{driveId}/items/{rootItemId}/delta` | `GET /drives/{driveId}/items/{rootItemId}/delta` |
| **Optimistic Concurrency Upload** | `PUT /drives/{driveId}/items/{rootItemId}:/{path}:/content` (`If-Match: {eTag}`) | Identical (`HTTP 412 Precondition Failed` on concurrent edit) |
| **Version History** | `GET /drives/{driveId}/items/{itemId}/versions` | Identical |
| **Version Restore / Download** | `GET /drives/{driveId}/items/{itemId}/versions/{versionId}/content` | Identical |
| **Per-Vault Sharing** | `POST /drives/{driveId}/items/{rootItemId}/invite` | Identical |
| **Item Deletion** | `DELETE /drives/{driveId}/items/{itemId}` | Identical |

---

## 2. Authentication & Token Lifecycle (`src/auth.ts`)

- **Authorization Code Flow with PKCE**: Uses `obsidian://entra-sync-auth` as a public client native redirect URI. No client secret is required or stored.
- **Token Refresh & Automatic 401 Recovery**: Access tokens are automatically refreshed before expiration using the `offline_access` refresh token. If Microsoft Graph returns `401 Unauthorized` mid-session, `GraphVaultClient` forces an immediate token refresh and retries the request once.
- **Download Fallback**: Pre-authenticated `@microsoft.graph.downloadUrl` links are tried first without an `Authorization` header; if a corporate proxy or personal OneDrive endpoint rejects the pre-authenticated URL, the client falls back to authenticated `/drives/{driveId}/items/{itemId}/content`.

---

## 3. Incremental Delta Sync & 3-Way Merge (`src/sync.ts`)

1. **Pull Phase (`/delta`)**:
   - Queries `/drives/{driveId}/items/{rootItemId}/delta` using the last persisted `@odata.deltaLink`.
   - If the cached `deltaLink` expires (`401` or `410 Gone`), the client automatically falls back to a full delta enumeration.
   - Remote deletions (`item.deleted`) remove the corresponding local file or folder.
   - If a remote text file changed while the local file remained unchanged since the last sync, the remote version is written directly.
   - If both the remote file and the local file changed since the last synced `baseText`, the engine runs a line-based **3-way `diff3` merge** (`node-diff3`):
     - Clean merges are written to disk and pushed back to SharePoint/OneDrive so all collaborators converge.
     - Overlapping line edits insert standard `<<<<<<< LOCAL` / `=======` / `>>>>>>> REMOTE` conflict markers.

2. **Local Deletion Reconciliation**:
   - Local `delete` and `rename` vault events propagate immediately to Microsoft Graph.
   - During full sync cycles, any previously tracked file or folder missing from the local vault (and not newly created on the remote) is deleted from the remote drive.

3. **Push Phase (`If-Match: {eTag}`)**:
   - Uploads modified text and binary files with the `If-Match: {eTag}` header.
   - If another user updated the remote file first, Microsoft Graph responds with `412 Precondition Failed`. The engine fetches the latest remote revision, executes a 3-way merge against `baseText`, and uploads the merged result with the new `eTag`.

---

## 4. Vault Configuration (`.obsidian`) Isolation

When **Sync vault configuration (`.obsidian`)** is enabled, shared settings (`app.json`, `appearance.json`, `core-plugins.json`, `community-plugins.json`, `hotkeys.json`, `types.json`, `bookmarks.json`, `snippets/**`, `themes/**`, and installed plugins) are synchronized via `Vault.adapter`.

The following paths are **always excluded** (`isIgnoredVaultPath`) so credentials and personal window layouts never leak across collaborators:
- `.obsidian/plugins/entra-sharepoint-sync/data.json` (contains the user's OAuth2 tokens and local sync state)
- `.obsidian/workspace.json` and `.obsidian/workspace-mobile.json` (per-user open panes and cursor state)
- `.DS_Store`, `.git/**`, and `.trash/**`

---

## 5. Automated Testing

- **Unit & Integration Suite (`test/sync.test.mjs`)**: Runs against a local stateful Microsoft Graph `/v1.0/drives` mock (`test/mock-graph-server.mjs`) that emulates `eTag` concurrency (`If-Match` / `412`), `/delta` tokens, `/versions`, `/invite`, binary attachments, bidirectional deletions, and `.obsidian` exclusions.
- **Two-User Docker E2E Suite (`test/docker-e2e.mjs`)**: Launches two `lscr.io/linuxserver/obsidian` containers with Playwright browser automation to validate multi-user sign-in, vault sharing, concurrent editing, and version history end to end.
