# Obsidian Enterprise Sync (Entra ID & SharePoint / OneDrive)

Self-hosted / zero-backend enterprise Obsidian synchronization and multi-user collaboration plugin powered by **Microsoft Entra ID** and **Microsoft Graph (SharePoint Online & OneDrive)**.

Signing in with an organizational **Microsoft Entra ID** (or personal Microsoft OneDrive) account unlocks an experience modeled after **OneNote on SharePoint**:
- **Enterprise Entra ID Sign-In**: OAuth 2.0 Authorization Code + PKCE (`obsidian://entra-sync-auth`), Device Code flow, and direct Entra ID credentials sign-in (`https://login.microsoftonline.com/common`).
- **Per-Vault Sharing & Permissions**: Create a vault folder on SharePoint or OneDrive (e.g., `/ObsidianVaults/{Vault}` or `/Documents/Documents/{Vault}`), invite colleagues by email (`POST /drives/{driveId}/items/{itemId}/invite`) with `write` or `read` permissions, and discover vaults shared with you (`/me/drive/sharedWithMe` and `/me/followedSites`).
- **Incremental Delta Sync & Automatic 3-Way Merge**: Uses Microsoft Graph `/delta` and optimistic concurrency (`PUT .../content` with `If-Match: {eTag}`). When two users edit the same note concurrently (`HTTP 412 Precondition Failed`), the plugin automatically performs a line-based **3-way `diff3` merge** against the last synced base revision without creating duplicate conflict files.
- **Per-Note Version History & Restore**: Inspect every historical version of a note (`GET /drives/{driveId}/items/{itemId}/versions`) with author display name, email, timestamp, content preview, and one-click restore.

---

## 1. Request for Global Tenant Admin (Nederlands)

Stuur onderstaand bericht naar jullie **Global / Entra ID Administrator**:

---

**Onderwerp:** Aanvraag Entra ID App Registration voor Obsidian Enterprise Sync (Microsoft Graph)

Hoi [Naam],

Voor het veilig samenwerken in Obsidian-vaults bovenop onze eigen SharePoint / OneDrive omgeving willen we een interne Obsidian-plugin koppelen via Entra ID (OAuth2 + PKCE, zonder client secret).

Zou je hiervoor een **App Registration** in Microsoft Entra ID willen aanmaken met de volgende instellingen?

1. **Naam:** `Obsidian Enterprise Sync`
2. **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts* (of *Single tenant* voor alleen onze organisatie)
3. **Authentication (Platform configuratie):**
   - Platform: **Mobile and desktop applications** (Public client / Native)
   - Custom Redirect URI: `obsidian://entra-sync-auth`
4. **API Permissions (Microsoft Graph — Delegated permissions):**
   - `User.Read` (inloggen en auteursnaam tonen)
   - `Files.ReadWrite.All` (lezen/schrijven van gedeelde vaults en versiegeschiedenis)
   - `Sites.ReadWrite.All` (optioneel: voor het uitlezen van SharePoint Document Libraries van teams)
   - `offline_access` (refresh tokens zodat gebruikers ingelogd blijven)
   - Graag **Grant admin consent for [Organisatie]** uitvoeren op deze gedelegeerde permissies.

**Wat ik daarna nodig heb:**
- De **Application (client) ID**
- De **Directory (tenant) ID**

---

## 2. Consumer OneDrive vs. Enterprise SharePoint Compatibility

Because the plugin addresses every vault via `/v1.0/drives/{driveId}/items/{rootItemId}`, **Consumer OneDrive** (`onedrive.live.com`, e.g., a folder inside `/Documents/Documents/...` or `/ObsidianVaults/...`) and **Enterprise SharePoint Document Libraries** use the exact same Microsoft Graph endpoints:

| Feature | Consumer OneDrive (`common`) | Enterprise SharePoint (`{tenantId}`) |
|---|---|---|
| **Auth Endpoint** | `login.microsoftonline.com/common` | `login.microsoftonline.com/{tenant-id}` |
| **Vault Container** | Folder in `/me/drive` & `/me/drive/sharedWithMe` | SharePoint Document Library (`/sites/{id}/drive`) or shared folder |
| **Delta Sync** | `GET /drives/{driveId}/items/{rootId}/delta` | Identical (`GET /drives/{driveId}/items/{rootId}/delta`) |
| **Upload + Conflict Check** | `PUT .../content` with `If-Match: {eTag}` (`412` on conflict) | Identical |
| **Version History** | `GET /drives/{driveId}/items/{itemId}/versions` | Identical |
| **Per-Vault Sharing** | `POST /drives/{driveId}/items/{rootId}/invite` | Identical |

---

## 3. Building & Testing

```bash
npm install
npm run build
npm test
```
