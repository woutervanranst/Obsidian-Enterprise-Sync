# OneDrive + Enterprise SharePoint + Teams Sync

Collaborate on Obsidian vaults with your team or across your personal devices using the **Microsoft OneDrive**, **SharePoint Online**, and **Microsoft Teams** storage you already have.

This plugin brings a **OneNote on SharePoint** workflow to Obsidian: sign in with your Microsoft account or corporate Entra ID, share individual vaults with colleagues, merge concurrent edits automatically without duplicate conflict files, and browse or restore previous note versions with full author attribution.

## Why

In Microsoft 365 organizations, OneNote notebooks live inside SharePoint and Microsoft Teams: access is governed by Entra ID, notebooks can be shared per project or team, and version history is built in.

For Obsidian, nothing comparable existed:
- **Official Obsidian Sync** requires a separate subscription and stores vault data outside your organization's Microsoft 365 tenant.
- **Syncing a vault folder with the OneDrive desktop client** frequently creates `Note (Work PC's conflicted copy).md` duplicates whenever two people edit a note around the same time, and breaks when syncing `.obsidian` configuration across machines.
- **Custom sync tools (CouchDB / LiveSync, Git, S3/WebDAV)** require provisioning external servers, managing credentials, or teaching non-technical teammates Git workflows.

**OneDrive + Enterprise SharePoint + Teams Sync** talks directly to the Microsoft Graph API from Obsidian—no extra servers to host, no external databases, and all data stays inside your Microsoft account or corporate tenant.

## Installation

Install directly from the official Obsidian Community Plugins directory:

1. Open **Settings → Community plugins** in Obsidian.
2. Select **Browse** and search for **OneDrive + Enterprise SharePoint + Teams Sync**.
3. Click **Install**, then **Enable**.

## Setup Wizard Overview

Open **Settings → OneDrive + Enterprise SharePoint + Teams Sync** to launch the guided 3-step onboarding wizard:

1. **Step 1 — Account type**: Choose **Consumer (OneDrive)** or **Enterprise (Teams & SharePoint)**.
2. **Step 2 — Sign in**: Authenticate in your system browser via OAuth 2.0 + PKCE (`obsidian://entra-sync-auth`). No passwords or credentials are ever entered inside Obsidian.
3. **Step 3 — Connect vault**: Choose **Discover vaults**, **Open vault by URL**, or **Create vault**. Once connected, the settings page switches to a clean **Connected Vault** view.

## Consumer (OneDrive)

You can use a personal Microsoft account (`outlook.com`, `live.com`, `hotmail.com`) to sync vaults across your devices or collaborate on a shared OneDrive folder with friends and colleagues.

1. Open **Settings → OneDrive + Enterprise SharePoint + Teams Sync**.
2. **Step 1 — Account type**: Select **Consumer (OneDrive)** and click **Continue**.
3. **Step 2 — Sign in**: Click **Sign in with browser** and sign in with your personal Microsoft account (`User.Read`, `Files.ReadWrite.All`, and `offline_access`).
4. **Step 3 — Connect vault**:
   - **Discover vaults**: Automatically lists vaults inside your OneDrive folder (default `/ObsidianVaults`—click **Change folder...** to customize, e.g. `Documents/Vaults`) as well as folders **Shared with me**, then click **Connect**.
   - **Open vault by URL**: Paste a shared OneDrive link (`https://1drv.ms/...` or `https://onedrive.live.com/...`) and click **Open vault**.
   - **Create vault**: Enter a **Vault folder name** (and optional **Parent folder in OneDrive**, default `ObsidianVaults`) and click **Create and connect vault**.
5. To invite collaborators, click **Share vault...** on the connected vault card, enter their email addresses, choose **Can edit (read and write)** or **View only (read)**, and click **Send invite**.

## Enterprise (Teams & SharePoint)

### One-Time Admin Setup

To use the plugin with corporate **Microsoft Entra ID** accounts, SharePoint Online, Microsoft Teams, and OneDrive for Business, a Global or Cloud Application Administrator must register a public client application in Microsoft Entra ID once for your organization.

> Tip: You can also view and copy this exact request directly from **Step 2** of the plugin's setup wizard (**Need these IDs? View one-time IT admin setup request**).

Send the following request to your IT / Entra ID administrator:

> **Subject:** Request for Microsoft Entra ID App Registration — Obsidian SharePoint & OneDrive Sync
>
> Hi [Name],
>
> We would like to use the **OneDrive + Enterprise SharePoint + Teams Sync** plugin in Obsidian to collaborate on team documentation stored in our SharePoint Online, Microsoft Teams, and OneDrive environment. The plugin connects directly to Microsoft Graph using OAuth 2.0 Authorization Code Flow with PKCE (public client, no client secret).
>
> Could you create an **App Registration** in **Microsoft Entra ID** with the following settings?
>
> 1. **Name:** `Obsidian SharePoint Sync`
> 2. **Supported account types:** *Accounts in this organizational directory only (Single tenant)* (or *Multitenant* if external guests also need access)
> 3. **Authentication (Platform configuration):**
>    - Platform: **Mobile and desktop applications**
>    - Custom Redirect URI: `obsidian://entra-sync-auth`
> 4. **API permissions (Microsoft Graph — Delegated permissions):**
>    - `User.Read` (display signed-in user profile and version authors)
>    - `Files.ReadWrite.All` (sync shared OneDrive and SharePoint vault files and version history)
>    - `Sites.ReadWrite.All` (discover and sync followed SharePoint sites and Microsoft Teams document libraries)
>    - `offline_access` (refresh tokens so users stay signed in)
>    - Please click **Grant admin consent for [Organization]** for these delegated permissions.
>
> Once created, please share back:
> - **Directory (tenant) ID**
> - **Application (client) ID**

### Signing In & Connecting an Enterprise Vault

1. Open **Settings → OneDrive + Enterprise SharePoint + Teams Sync**.
2. **Step 1 — Account type**: Select **Enterprise (Teams & SharePoint)** and click **Continue**. SharePoint and Microsoft Teams access (`Sites.ReadWrite.All`) is automatically enabled for Enterprise accounts.
3. **Step 2 — Sign in**: Enter your organization's **Directory (tenant) ID** and **Application (client) ID**, then click **Sign in with browser**.
4. **Step 3 — Connect vault**:
   - **Discover vaults**: Browse vaults across your followed SharePoint & Microsoft Teams sites, **Shared with me** folders, and your **OneDrive for Business**, then click **Connect**.
   - **Open vault by URL**:
     - **SharePoint Online**: Open the target SharePoint Document Library or subfolder in your browser, copy the address bar URL (including `Forms/AllItems.aspx?id=...` or a direct folder link), paste it into **Vault folder URL**, and click **Open vault**.
     - **Microsoft Teams**: Open your Teams channel → **Files** tab → click **Open in SharePoint**, copy the browser URL, paste it into **Vault folder URL**, and click **Open vault**.
   - **Create vault**: Choose **SharePoint or Teams library** (paste the parent library/channel URL and enter a new **Vault folder name**) or **My OneDrive for Business** (choose a parent folder in OneDrive and enter a **Vault folder name**), then click **Create and connect vault**.

## Features

- **Microsoft Entra ID & Personal Microsoft Account Sign-In**: OAuth 2.0 + PKCE (`obsidian://entra-sync-auth`) with automatic background token refresh.
- **Per-Vault Sharing & Access Control**: Invite colleagues with read/write or read-only permissions directly from Obsidian (`Share vault with colleagues`) and view active collaborators.
- **Automatic 3-Way Conflict Merge (`diff3`)**: Uses SharePoint/OneDrive `eTag` optimistic concurrency (`If-Match`). When two collaborators edit different sections of the same note simultaneously, changes are automatically 3-way merged without creating duplicate conflict files.
- **Native Version History & Restore**: Inspect every historical revision of a note stored by SharePoint/OneDrive—including author name, email, timestamp, and full content preview—and restore any prior version in one click.
- **Attachments & Bidirectional Deletions**: Syncs Markdown notes, Canvas files, and binary attachments (`.png`, `.jpg`, `.pdf`, audio/video, etc.), and propagates file and folder renames and deletions in both directions.
- **Safe Shared Vault Configuration (`.obsidian`)**: Optionally sync shared themes, CSS snippets, hotkeys, and plugins across your team while strictly excluding your personal OAuth tokens (`data.json`) and window/tab layout (`workspace.json`).

## Security & Privacy

- **Direct-to-Microsoft Communication**: All authentication and file requests go directly between your Obsidian app and Microsoft (`https://login.microsoftonline.com` and `https://graph.microsoft.com`). There are no third-party relay servers, analytics, or telemetry.
- **PKCE Public Client Flow**: Authentication uses OAuth 2.0 Authorization Code with PKCE; no client secrets are embedded in the plugin.
- **Credential Isolation**: Your local `plugins/entra-sharepoint-sync/data.json` file (which stores your OAuth tokens and local sync cursor) is permanently excluded from `.obsidian` vault synchronization so credentials are never uploaded to a shared drive.

## Troubleshooting

If you encounter a sync issue or authentication error:

1. Open **Settings → OneDrive + Enterprise SharePoint + Teams Sync** and expand **Troubleshooting and diagnostic logs** at the bottom of the page (or run the command **View sync diagnostic logs** from the Command Palette).
2. Click **Force token refresh** to verify your OAuth2 refresh token and obtain a fresh access token.
3. Click **Reset delta and resync** if a folder was reorganized externally in SharePoint/OneDrive and you want to force a full delta reconciliation.
4. Use **Copy logs** to copy sanitized diagnostic entries (tokens and signatures are automatically redacted) when reporting issues.

## Building & Testing

### Unit & Integration Tests

```bash
npm install
npm run build
npm test
```

### Multi-User Docker E2E Testing (`linuxserver/docker-obsidian`)

End-to-end multi-user collaboration is tested using two isolated [`linuxserver/docker-obsidian`](https://github.com/linuxserver/docker-obsidian) (`lscr.io/linuxserver/obsidian:latest`) containers and Playwright browser automation:

```bash
npm run test:docker
```

`test/docker-e2e.mjs` performs the following steps automatically:
1. Starts two `linuxserver/docker-obsidian` containers (`obsidian-entra-alice` and `obsidian-entra-bob`), each running a full desktop Obsidian instance accessible over its web UI with the built plugin (`main.js`, `manifest.json`, `styles.css`) mounted into each user's vault.
2. Uses Playwright to sign both containers in with two distinct Microsoft Entra ID test users.
3. Has User 1 create a shared vault, invite User 2 via the in-app **Share vault** modal (`ShareVaultModal`), and push initial notes.
4. Has User 2 discover the shared vault, pull the notes via `/delta`, and make concurrent edits alongside User 1 to verify `HTTP 412 Precondition Failed` handling and automatic **3-way `diff3` conflict merging**.
5. Opens the **Version history** modal (`VersionHistoryModal`) in Obsidian to verify revision attribution and one-click version restore.

For deep technical details on the Microsoft Graph `/drives` mapping, `eTag` + `diff3` concurrency model, and `.obsidian` exclusion rules, see [ARCHITECTURE.md](./ARCHITECTURE.md).

## License

[MIT](./LICENSE)

