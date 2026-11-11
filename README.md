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

## Consumer

### OneDrive

You can use a personal Microsoft account (`outlook.com`, `live.com`, `hotmail.com`) to sync vaults across your devices or collaborate on a shared OneDrive folder with friends and colleagues.

1. Open **Settings → OneDrive + Enterprise SharePoint + Teams Sync**.
2. Leave **Tenant ID** set to `common`.
3. Optional: Set **OneDrive vaults folder path** to the parent folder in your OneDrive where your vaults live (default: `ObsidianVaults`, or e.g. `Documents/Vaults`).
4. Click **Sign in (browser PKCE)** and sign in with your personal Microsoft account.
5. Under **Shared vault and collaboration**:
   - **Create a new vault**: Enter a folder name in **Create or connect vault folder** and click **Create and connect vault**.
   - **Connect an existing or shared vault**: Click **Discover vaults** and pick a folder from your OneDrive or **Shared with me** in the dropdown.
6. To invite collaborators, click **Share vault...**, enter their email addresses, choose **Can edit (read and write)** or **View only (read)**, and click **Send invite**.

## Enterprise

### One-Time Admin Setup

To use the plugin with corporate **Microsoft Entra ID** accounts, SharePoint Online, and Microsoft Teams, a Global or Cloud Application Administrator must register a public client application in Microsoft Entra ID once for your organization.

Send the following request to your IT / Entra ID administrator:

> **Subject:** Request for Microsoft Entra ID App Registration — Obsidian SharePoint & OneDrive Sync
>
> Hi [Name],
>
> We would like to use the **OneDrive + Enterprise SharePoint + Teams Sync** plugin in Obsidian to collaborate on team documentation stored in our SharePoint Online / OneDrive environment. The plugin connects directly to Microsoft Graph using OAuth 2.0 Authorization Code Flow with PKCE (public client, no client secret).
>
> Could you create an **App Registration** in **Microsoft Entra ID** with the following settings?
>
> 1. **Name:** `Obsidian SharePoint Sync`
> 2. **Supported account types:** *Accounts in this organizational directory only (Single tenant)* (or *Multitenant* if external guests also need access)
> 3. **Authentication (Platform configuration):**
>    - Platform: **Mobile and desktop applications**
>    - Custom Redirect URI: `obsidian://entra-sync-auth`
>    - Enable **Allow public client flows**: *Yes* (optional, only needed if using device-code or direct credential flows)
> 4. **API permissions (Microsoft Graph — Delegated permissions):**
>    - `User.Read` (display signed-in user profile and version authors)
>    - `Files.ReadWrite.All` (sync shared OneDrive and SharePoint vault files and version history)
>    - `Sites.ReadWrite.All` (discover and sync followed SharePoint sites and Teams document libraries)
>    - `offline_access` (refresh tokens so users stay signed in)
>    - Please click **Grant admin consent for [Organization]** for these delegated permissions.
>
> Once created, please share back:
> - **Application (client) ID**
> - **Directory (tenant) ID**

Once you receive those two IDs from your admin, paste them into **Tenant ID** and **Entra app client ID** in the plugin settings and click **Sign in (browser PKCE)**.

### SharePoint Configuration

You can use a SharePoint Site Document Library (or a shared folder inside it) as a collaborative team vault:

1. Open the target SharePoint site in your browser and click the **Follow** star (top right) so Microsoft Graph lists it under your followed sites—or have a teammate share a specific folder with you.
2. In Obsidian, open **Settings → OneDrive + Enterprise SharePoint + Teams Sync**.
3. Sign in with your corporate Entra ID account.
4. Click **Discover vaults** and select either:
   - Your followed SharePoint site (`<Site Name> (SharePoint)`), or
   - A specific shared folder (`<Folder Name> (Shared with me)`).
5. Click **Sync now** (or let automatic background sync run).

### Microsoft Teams Configuration

Every Microsoft Teams team and standard channel is backed by a SharePoint Online site and Document Library (`Files` tab in Teams):

1. In Microsoft Teams, open the channel where you want to store the team's Obsidian vault, go to the **Files** tab, and click **Open in SharePoint**.
2. Choose one of two approaches:
   - **Share a dedicated vault folder (Recommended)**: Create a folder inside the channel's `Files` tab (e.g., `EngineeringVault`), click **Share** on that folder for your team members, and select `<Folder Name> (Shared with me)` after clicking **Discover vaults** in Obsidian.
   - **Follow the Team's SharePoint site**: Click **Follow** (star icon) on the Team's SharePoint page in your browser, then click **Discover vaults** in Obsidian to select `<Team Name> (SharePoint)`.

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

1. Open **Settings → OneDrive + Enterprise SharePoint + Teams Sync** and scroll down to **Troubleshooting and sync logs** (or run the command **View sync diagnostic logs** from the Command Palette).
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

