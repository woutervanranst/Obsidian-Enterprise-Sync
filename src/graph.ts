import { requestUrl } from "obsidian";

export interface RemoteVaultTarget {
  name: string;
  driveId: string;
  rootItemId: string;
  webUrl: string;
}

export interface DeltaDriveItem {
  id: string;
  name: string;
  eTag?: string;
  cTag?: string;
  deleted?: { state: string };
  folder?: { childCount: number };
  file?: { mimeType: string };
  parentReference?: { id: string; driveId?: string; path?: string };
  lastModifiedDateTime?: string;
  lastModifiedBy?: {
    user?: { displayName?: string; email?: string };
  };
}

export interface DriveItemVersion {
  id: string;
  lastModifiedDateTime: string;
  lastModifiedBy?: {
    user?: { displayName?: string; email?: string };
  };
  size?: number;
}

export interface VaultPermission {
  id: string;
  roles: string[];
  grantedToV2?: {
    user?: { displayName?: string; email?: string };
  };
  invitation?: {
    email?: string;
  };
}

/**
 * Unified Microsoft Graph client that works identically for:
 *  - Consumer OneDrive shared folders (/drives/{driveId}/items/{folderId})
 *    including custom paths like /Documents/Documents/{vault} or /ObsidianVaults/{vault}
 *  - Enterprise SharePoint Document Libraries (/drives/{driveId}/items/root or subfolder)
 */
export class GraphVaultClient {
  constructor(
    private getAccessToken: () => Promise<string>,
    private getBaseUrl: () => string = () => "https://graph.microsoft.com/v1.0"
  ) {}

  private baseUrl(): string {
    return this.getBaseUrl().replace(/\/+$/, "");
  }

  private async headers(extra: Record<string, string> = {}): Promise<Record<string, string>> {
    const token = await this.getAccessToken();
    return {
      Authorization: `Bearer ${token}`,
      ...extra,
    };
  }

  /**
   * Discover available vaults from:
   * 1. User's own OneDrive (/ObsidianVaults/* and optionally a custom parent folder like /Documents/Documents)
   * 2. Folders shared with the user (/me/drive/sharedWithMe)
   * 3. Enterprise SharePoint sites the user follows (/me/followedSites -> /sites/{id}/drive)
   */
  async listAvailableVaults(
    includeSharePointSites = false,
    customOneDriveFolder = "ObsidianVaults"
  ): Promise<RemoteVaultTarget[]> {
    const vaults: RemoteVaultTarget[] = [];
    const seen = new Set<string>();
    const headers = await this.headers();
    const base = this.baseUrl();

    const addVault = (v: RemoteVaultTarget) => {
      const key = `${v.driveId}:${v.rootItemId}`;
      if (!seen.has(key)) {
        seen.add(key);
        vaults.push(v);
      }
    };

    const foldersToCheck = Array.from(
      new Set(["ObsidianVaults", customOneDriveFolder.replace(/^\/+|\/+$/g, "")].filter(Boolean))
    );

    // 1. User's own OneDrive folders
    for (const folderPath of foldersToCheck) {
      const encodedFolder = folderPath
        .split("/")
        .map((s) => encodeURIComponent(s))
        .join("/");
      const ownFolderRes = await requestUrl({
        url: `${base}/me/drive/root:/${encodedFolder}:/children`,
        headers,
        throw: false,
      });
      if (ownFolderRes.status === 200) {
        for (const item of ownFolderRes.json.value ?? []) {
          if (item.folder && item.parentReference?.driveId) {
            addVault({
              name: `${item.name} (My OneDrive)`,
              driveId: item.parentReference.driveId,
              rootItemId: item.id,
              webUrl: item.webUrl ?? "",
            });
          }
        }
      }
    }

    // 2. Folders shared with me (works in Consumer OneDrive AND OneDrive for Business / SharePoint)
    const sharedRes = await requestUrl({
      url: `${base}/me/drive/sharedWithMe`,
      headers,
      throw: false,
    });
    if (sharedRes.status === 200) {
      for (const item of sharedRes.json.value ?? []) {
        const remote = item.remoteItem ?? item;
        if (remote.folder && remote.parentReference?.driveId) {
          addVault({
            name: `${remote.name} (Shared with me)`,
            driveId: remote.parentReference.driveId,
            rootItemId: remote.id,
            webUrl: remote.webUrl ?? item.webUrl ?? "",
          });
        }
      }
    }

    // 3. Enterprise SharePoint sites user follows (Business / Enterprise only)
    if (includeSharePointSites) {
      const sitesRes = await requestUrl({
        url: `${base}/me/followedSites`,
        headers,
        throw: false,
      });
      if (sitesRes.status === 200) {
        for (const site of sitesRes.json.value ?? []) {
          const driveRes = await requestUrl({
            url: `${base}/sites/${site.id}/drive`,
            headers,
            throw: false,
          });
          if (driveRes.status === 200) {
            addVault({
              name: `${site.displayName} (SharePoint)`,
              driveId: driveRes.json.id,
              rootItemId: "root",
              webUrl: driveRes.json.webUrl ?? "",
            });
          }
        }
      }
    }

    return vaults;
  }

  /**
   * Create or resolve a vault folder in the user's OneDrive/SharePoint drive by path
   * (e.g., "ObsidianVaults/EngineeringWiki" or "Documents/Documents/EngineeringWiki").
   */
  async createOrGetVaultFolder(
    vaultName: string,
    parentFolder = "ObsidianVaults"
  ): Promise<RemoteVaultTarget> {
    const headers = await this.headers({ "Content-Type": "application/json" });
    const base = this.baseUrl();
    const cleanParent = parentFolder.replace(/^\/+|\/+$/g, "");
    const cleanName = vaultName.replace(/^\/+|\/+$/g, "");
    const fullPath = cleanParent ? `${cleanParent}/${cleanName}` : cleanName;

    // Ensure each segment exists on /me/drive
    const segments = fullPath.split("/").filter(Boolean);
    let currentParentId = "root";
    let driveId = "";
    let lastItem: any = null;

    for (const segment of segments) {
      const createUrl =
        currentParentId === "root"
          ? `${base}/me/drive/root/children`
          : `${base}/me/drive/items/${currentParentId}/children`;

      const res = await requestUrl({
        url: createUrl,
        method: "POST",
        headers,
        body: JSON.stringify({
          name: segment,
          folder: {},
          "@microsoft.graph.conflictBehavior": "replace",
        }),
        throw: false,
      });

      if (res.status >= 200 && res.status < 300 && res.json) {
        lastItem = res.json;
        currentParentId = lastItem.id;
        driveId = lastItem.parentReference?.driveId ?? driveId;
      } else {
        // Fallback: look up by path
        const subPath = segments
          .slice(0, segments.indexOf(segment) + 1)
          .map((s) => encodeURIComponent(s))
          .join("/");
        const getRes = await requestUrl({
          url: `${base}/me/drive/root:/${subPath}`,
          headers: await this.headers(),
        });
        lastItem = getRes.json;
        currentParentId = lastItem.id;
        driveId = lastItem.parentReference?.driveId ?? driveId;
      }
    }

    return {
      name: `${cleanName} (My OneDrive)`,
      driveId,
      rootItemId: currentParentId,
      webUrl: lastItem?.webUrl ?? "",
    };
  }

  /**
   * Share a vault folder with colleagues by email (works on Consumer OneDrive & SharePoint)
   */
  async inviteCollaborators(
    target: RemoteVaultTarget,
    emails: string[],
    role: "write" | "read" = "write"
  ): Promise<VaultPermission[]> {
    const headers = await this.headers({ "Content-Type": "application/json" });
    const base = this.baseUrl();
    const res = await requestUrl({
      url: `${base}/drives/${target.driveId}/items/${target.rootItemId}/invite`,
      method: "POST",
      headers,
      body: JSON.stringify({
        requireSignIn: true,
        sendInvitation: true,
        roles: [role],
        recipients: emails.map((email) => ({ email: email.trim() })),
      }),
    });
    return res.json?.value ?? [];
  }

  async listPermissions(target: RemoteVaultTarget): Promise<VaultPermission[]> {
    const headers = await this.headers();
    const base = this.baseUrl();
    const res = await requestUrl({
      url: `${base}/drives/${target.driveId}/items/${target.rootItemId}/permissions`,
      headers,
      throw: false,
    });
    if (res.status !== 200) return [];
    return res.json?.value ?? [];
  }

  /**
   * Fetch incremental changes since the last deltaLink
   */
  async fetchDelta(
    target: RemoteVaultTarget,
    deltaLink?: string
  ): Promise<{ items: DeltaDriveItem[]; nextDeltaLink: string }> {
    const headers = await this.headers();
    const base = this.baseUrl();
    let url =
      deltaLink || `${base}/drives/${target.driveId}/items/${target.rootItemId}/delta`;

    const items: DeltaDriveItem[] = [];
    let nextDeltaLink = "";

    while (url) {
      const res = await requestUrl({ url, headers });
      items.push(...(res.json.value ?? []));
      if (res.json["@odata.nextLink"]) {
        url = res.json["@odata.nextLink"];
      } else {
        nextDeltaLink = res.json["@odata.deltaLink"] ?? "";
        break;
      }
    }

    return { items, nextDeltaLink };
  }

  async downloadFileText(
    driveId: string,
    itemId: string
  ): Promise<{ text: string; eTag: string; lastModifiedBy?: string }> {
    const headers = await this.headers();
    const base = this.baseUrl();
    const metaRes = await requestUrl({
      url: `${base}/drives/${driveId}/items/${itemId}`,
      headers,
    });
    const downloadUrl =
      metaRes.json["@microsoft.graph.downloadUrl"] ||
      `${base}/drives/${driveId}/items/${itemId}/content`;

    const contentRes = await requestUrl({
      url: downloadUrl,
      headers: downloadUrl.startsWith(base) ? headers : undefined,
    });
    const author =
      metaRes.json.lastModifiedBy?.user?.displayName ||
      metaRes.json.lastModifiedBy?.user?.email;

    return {
      text: contentRes.text,
      eTag: metaRes.json.eTag ?? "",
      lastModifiedBy: author,
    };
  }

  /**
   * Upload file content with optimistic concurrency (If-Match: eTag).
   * Returns 412 status if another collaborator modified the file concurrently.
   */
  async uploadFileText(
    target: RemoteVaultTarget,
    relativePath: string,
    content: string,
    expectedETag?: string
  ): Promise<{ status: number; item?: DeltaDriveItem }> {
    const extra: Record<string, string> = { "Content-Type": "text/plain; charset=utf-8" };
    if (expectedETag) {
      extra["If-Match"] = expectedETag;
    }
    const headers = await this.headers(extra);
    const base = this.baseUrl();
    const encodedPath = relativePath
      .split("/")
      .map((seg) => encodeURIComponent(seg))
      .join("/");

    const res = await requestUrl({
      url: `${base}/drives/${target.driveId}/items/${target.rootItemId}:/${encodedPath}:/content`,
      method: "PUT",
      headers,
      body: content,
      throw: false,
    });

    return {
      status: res.status,
      item: res.status >= 200 && res.status < 300 ? res.json : undefined,
    };
  }

  /**
   * List version history for a note (author + timestamp)
   */
  async listVersions(driveId: string, itemId: string): Promise<DriveItemVersion[]> {
    const headers = await this.headers();
    const base = this.baseUrl();
    const res = await requestUrl({
      url: `${base}/drives/${driveId}/items/${itemId}/versions`,
      headers,
    });
    return res.json.value ?? [];
  }

  async downloadVersionText(
    driveId: string,
    itemId: string,
    versionId: string
  ): Promise<string> {
    const headers = await this.headers();
    const base = this.baseUrl();
    const res = await requestUrl({
      url: `${base}/drives/${driveId}/items/${itemId}/versions/${versionId}/content`,
      headers,
    });
    return res.text;
  }
}
