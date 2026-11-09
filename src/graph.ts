import { requestUrl, RequestUrlResponse } from "obsidian";
import { SyncLogger, defaultLogger } from "./logger";

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

interface GraphRequestOptions {
  url: string;
  method?: string;
  extraHeaders?: Record<string, string>;
  body?: string;
  throwOnError?: boolean;
  omitAuth?: boolean;
}

/**
 * Unified Microsoft Graph client that works identically for:
 *  - Consumer OneDrive shared folders (/drives/{driveId}/items/{folderId})
 *    including custom paths like /Documents/Documents/{vault} or /ObsidianVaults/{vault}
 *  - Enterprise SharePoint Document Libraries (/drives/{driveId}/items/root or subfolder)
 */
export class GraphVaultClient {
  constructor(
    private getAccessToken: (forceRefresh?: boolean) => Promise<string>,
    private getBaseUrl: () => string = () => "https://graph.microsoft.com/v1.0",
    private logger: SyncLogger = defaultLogger
  ) {}

  private baseUrl(): string {
    return this.getBaseUrl().replace(/\/+$/, "");
  }

  private sanitizeUrl(rawUrl: string): string {
    try {
      const parsed = new URL(rawUrl);
      for (const key of ["tempauth", "access_token", "sig", "token", "$skiptoken"]) {
        if (parsed.searchParams.has(key)) {
          const val = parsed.searchParams.get(key) || "";
          parsed.searchParams.set(key, `${val.slice(0, 6)}...`);
        }
      }
      return parsed.toString();
    } catch {
      return rawUrl;
    }
  }

  private extractErrorDetails(res: RequestUrlResponse): Record<string, unknown> {
    const wwwAuth =
      res.headers?.["www-authenticate"] || res.headers?.["WWW-Authenticate"] || undefined;
    const graphErr = res.json?.error;
    return {
      status: res.status,
      code: graphErr?.code,
      message: graphErr?.message,
      innerError: graphErr?.innerError,
      wwwAuthenticate: wwwAuth,
      bodyPreview: !graphErr && res.text ? res.text.slice(0, 300) : undefined,
    };
  }

  private formatErrorSummary(res: RequestUrlResponse): string {
    const graphErr = res.json?.error;
    if (graphErr?.code || graphErr?.message) {
      return `${graphErr.code || "GraphError"}: ${graphErr.message || ""}`.trim();
    }
    const wwwAuth =
      res.headers?.["www-authenticate"] || res.headers?.["WWW-Authenticate"];
    if (wwwAuth) {
      return `HTTP ${res.status} (${wwwAuth})`;
    }
    if (res.text) {
      return `HTTP ${res.status}: ${res.text.slice(0, 200)}`;
    }
    return `HTTP ${res.status}`;
  }

  private async headers(
    extra: Record<string, string> = {},
    forceRefresh = false
  ): Promise<Record<string, string>> {
    const token = await this.getAccessToken(forceRefresh);
    return {
      Authorization: `Bearer ${token}`,
      ...extra,
    };
  }

  /**
   * Executes an HTTP request via Obsidian's requestUrl with `throw: false`,
   * automatic 401 token refresh + single retry, and detailed diagnostic logging.
   */
  private async graphRequest(opts: GraphRequestOptions): Promise<RequestUrlResponse> {
    const method = opts.method || "GET";
    const safeUrl = this.sanitizeUrl(opts.url);
    const throwOnError = opts.throwOnError ?? true;

    let reqHeaders = opts.omitAuth
      ? { ...(opts.extraHeaders || {}) }
      : await this.headers(opts.extraHeaders || {}, false);

    let res = await requestUrl({
      url: opts.url,
      method,
      headers: reqHeaders,
      body: opts.body,
      throw: false,
    });

    if (res.status === 401 && !opts.omitAuth) {
      this.logger.warn(
        `${method} ${safeUrl} returned 401 Unauthorized — forcing OAuth2 token refresh and retrying once`,
        this.extractErrorDetails(res)
      );
      try {
        reqHeaders = await this.headers(opts.extraHeaders || {}, true);
        res = await requestUrl({
          url: opts.url,
          method,
          headers: reqHeaders,
          body: opts.body,
          throw: false,
        });
      } catch (refreshErr) {
        this.logger.error(
          `Token refresh failed after 401 on ${method} ${safeUrl}`,
          String(refreshErr)
        );
        if (throwOnError) {
          throw refreshErr;
        }
        return res;
      }
    }

    if (res.status >= 200 && res.status < 300) {
      this.logger.info(`${method} ${safeUrl} -> ${res.status}`);
    } else if (res.status === 412) {
      this.logger.warn(
        `${method} ${safeUrl} -> 412 Precondition Failed (concurrent edit detected)`
      );
    } else {
      const details = this.extractErrorDetails(res);
      if (throwOnError) {
        this.logger.error(`${method} ${safeUrl} -> ${res.status}`, details);
        throw new Error(
          `Graph ${method} ${safeUrl} failed (${res.status}): ${this.formatErrorSummary(res)}`
        );
      } else {
        this.logger.warn(`${method} ${safeUrl} -> ${res.status} (non-fatal)`, details);
      }
    }

    return res;
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
      const ownFolderRes = await this.graphRequest({
        url: `${base}/me/drive/root:/${encodedFolder}:/children`,
        throwOnError: false,
      });
      if (ownFolderRes.status === 200) {
        for (const item of ownFolderRes.json?.value ?? []) {
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
    const sharedRes = await this.graphRequest({
      url: `${base}/me/drive/sharedWithMe`,
      throwOnError: false,
    });
    if (sharedRes.status === 200) {
      for (const item of sharedRes.json?.value ?? []) {
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
      const sitesRes = await this.graphRequest({
        url: `${base}/me/followedSites`,
        throwOnError: false,
      });
      if (sitesRes.status === 200) {
        for (const site of sitesRes.json?.value ?? []) {
          const driveRes = await this.graphRequest({
            url: `${base}/sites/${site.id}/drive`,
            throwOnError: false,
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
    const base = this.baseUrl();
    const cleanParent = parentFolder.replace(/^\/+|\/+$/g, "");
    const cleanName = vaultName.replace(/^\/+|\/+$/g, "");
    const fullPath = cleanParent ? `${cleanParent}/${cleanName}` : cleanName;

    // Ensure each segment exists on /me/drive
    const segments = fullPath.split("/").filter(Boolean);
    let currentParentId = "root";
    let driveId = "";
    let lastItem: any = null;

    for (let i = 0; i < segments.length; i++) {
      const segment = segments[i];
      const createUrl =
        currentParentId === "root"
          ? `${base}/me/drive/root/children`
          : `${base}/me/drive/items/${currentParentId}/children`;

      const res = await this.graphRequest({
        url: createUrl,
        method: "POST",
        extraHeaders: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: segment,
          folder: {},
          "@microsoft.graph.conflictBehavior": "replace",
        }),
        throwOnError: false,
      });

      if (res.status >= 200 && res.status < 300 && res.json) {
        lastItem = res.json;
        currentParentId = lastItem.id;
        driveId = lastItem.parentReference?.driveId ?? driveId;
      } else {
        // Fallback: look up by path
        const subPath = segments
          .slice(0, i + 1)
          .map((s) => encodeURIComponent(s))
          .join("/");
        const getRes = await this.graphRequest({
          url: `${base}/me/drive/root:/${subPath}`,
          throwOnError: true,
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
    const base = this.baseUrl();
    const res = await this.graphRequest({
      url: `${base}/drives/${target.driveId}/items/${target.rootItemId}/invite`,
      method: "POST",
      extraHeaders: { "Content-Type": "application/json" },
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
    const base = this.baseUrl();
    const res = await this.graphRequest({
      url: `${base}/drives/${target.driveId}/items/${target.rootItemId}/permissions`,
      throwOnError: false,
    });
    if (res.status !== 200) return [];
    return res.json?.value ?? [];
  }

  /**
   * Fetch incremental changes since the last deltaLink.
   * If a cached deltaLink returns 401/410 (e.g. expired skipToken on OneDrive),
   * automatically resets to a fresh /delta query.
   */
  async fetchDelta(
    target: RemoteVaultTarget,
    deltaLink?: string
  ): Promise<{ items: DeltaDriveItem[]; nextDeltaLink: string }> {
    const base = this.baseUrl();
    const freshDeltaUrl = `${base}/drives/${target.driveId}/items/${target.rootItemId}/delta`;
    let url = deltaLink || freshDeltaUrl;

    const items: DeltaDriveItem[] = [];
    let nextDeltaLink = "";

    while (url) {
      let res = await this.graphRequest({
        url,
        throwOnError: url === freshDeltaUrl,
      });

      if ((res.status === 401 || res.status === 410) && url !== freshDeltaUrl) {
        this.logger.warn(
          `Cached deltaLink returned ${res.status}; falling back to fresh delta query at ${freshDeltaUrl}`
        );
        url = freshDeltaUrl;
        res = await this.graphRequest({ url, throwOnError: true });
      } else if (res.status < 200 || res.status >= 300) {
        throw new Error(
          `Graph GET ${this.sanitizeUrl(url)} failed (${res.status}): ${this.formatErrorSummary(res)}`
        );
      }

      items.push(...(res.json?.value ?? []));
      if (res.json?.["@odata.nextLink"]) {
        url = res.json["@odata.nextLink"];
      } else {
        nextDeltaLink = res.json?.["@odata.deltaLink"] ?? "";
        break;
      }
    }

    return { items, nextDeltaLink };
  }

  async downloadFileText(
    driveId: string,
    itemId: string
  ): Promise<{ text: string; eTag: string; lastModifiedBy?: string }> {
    const base = this.baseUrl();
    const metaRes = await this.graphRequest({
      url: `${base}/drives/${driveId}/items/${itemId}`,
    });
    const downloadUrl: string =
      metaRes.json?.["@microsoft.graph.downloadUrl"] ||
      `${base}/drives/${driveId}/items/${itemId}/content`;

    const isGraphApiUrl = downloadUrl.startsWith(base);

    // 1. Try pre-authenticated @microsoft.graph.downloadUrl without Bearer header first (unless it's a Graph API URL)
    let contentRes = await this.graphRequest({
      url: downloadUrl,
      omitAuth: !isGraphApiUrl,
      throwOnError: false,
    });

    // 2. If pre-authenticated URL returned 401/403 (or failed), retry with Bearer token
    if ((contentRes.status < 200 || contentRes.status >= 300) && !isGraphApiUrl) {
      this.logger.warn(
        `Pre-authenticated downloadUrl returned ${contentRes.status}, retrying with Bearer token`
      );
      contentRes = await this.graphRequest({
        url: downloadUrl,
        omitAuth: false,
        throwOnError: false,
      });
    }

    // 3. Final fallback: fetch directly from Graph /drives/{driveId}/items/{itemId}/content
    const graphContentUrl = `${base}/drives/${driveId}/items/${itemId}/content`;
    if (
      (contentRes.status < 200 || contentRes.status >= 300) &&
      downloadUrl !== graphContentUrl
    ) {
      this.logger.warn(
        `downloadUrl failed (${contentRes.status}), falling back to ${graphContentUrl}`
      );
      contentRes = await this.graphRequest({
        url: graphContentUrl,
        omitAuth: false,
        throwOnError: true,
      });
    } else if (contentRes.status < 200 || contentRes.status >= 300) {
      throw new Error(
        `File download failed (${contentRes.status}): ${this.formatErrorSummary(contentRes)}`
      );
    }

    const author =
      metaRes.json?.lastModifiedBy?.user?.displayName ||
      metaRes.json?.lastModifiedBy?.user?.email;

    return {
      text: contentRes.text,
      eTag: metaRes.json?.eTag ?? "",
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
    const base = this.baseUrl();
    const encodedPath = relativePath
      .split("/")
      .map((seg) => encodeURIComponent(seg))
      .join("/");

    const url = `${base}/drives/${target.driveId}/items/${target.rootItemId}:/${encodedPath}:/content`;
    const res = await this.graphRequest({
      url,
      method: "PUT",
      extraHeaders: extra,
      body: content,
      throwOnError: false,
    });

    if (res.status !== 412 && (res.status < 200 || res.status >= 300)) {
      throw new Error(
        `Graph PUT ${this.sanitizeUrl(url)} failed (${res.status}): ${this.formatErrorSummary(res)}`
      );
    }

    return {
      status: res.status,
      item: res.status >= 200 && res.status < 300 ? res.json : undefined,
    };
  }

  /**
   * List version history for a note (author + timestamp)
   */
  async listVersions(driveId: string, itemId: string): Promise<DriveItemVersion[]> {
    const base = this.baseUrl();
    const res = await this.graphRequest({
      url: `${base}/drives/${driveId}/items/${itemId}/versions`,
    });
    return res.json?.value ?? [];
  }

  async downloadVersionText(
    driveId: string,
    itemId: string,
    versionId: string
  ): Promise<string> {
    const base = this.baseUrl();
    const res = await this.graphRequest({
      url: `${base}/drives/${driveId}/items/${itemId}/versions/${versionId}/content`,
    });
    return res.text;
  }
}
