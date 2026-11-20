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
  name?: string;
  eTag?: string;
  cTag?: string;
  deleted?: { state: string };
  folder?: { childCount?: number };
  file?: { mimeType?: string; hashes?: { quickXorHash?: string } };
  parentReference?: { id: string; driveId?: string; path?: string; name?: string; driveType?: string; siteId?: string };
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
    user?: { displayName?: string; email?: string; id?: string };
    siteUser?: { displayName?: string; email?: string; id?: string; loginName?: string };
  };
  grantedTo?: {
    user?: { displayName?: string; email?: string; id?: string };
  };
  invitation?: {
    email?: string;
  };
  link?: {
    webUrl?: string;
    type?: string;
    scope?: string;
  };
}

interface GraphErrorPayload {
  error?: {
    code?: string;
    message?: string;
    innerError?: unknown;
  };
}

interface GraphDriveItemNode extends DeltaDriveItem {
  webUrl?: string;
  "@microsoft.graph.downloadUrl"?: string;
  remoteItem?: GraphDriveItemNode;
}

interface GraphListResponse<T> {
  value?: T[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

interface GraphSiteNode {
  id: string;
  displayName: string;
  webUrl?: string;
}

interface GraphDriveNode {
  id: string;
  name?: string;
  webUrl?: string;
}

export function encodeGraphShareToken(targetUrl: string): string {
  const bytes = new TextEncoder().encode(targetUrl);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  const base64Url = btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
  return `u!${base64Url}`;
}

function safeDecodeURIComponent(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

interface GraphRequestOptions {
  url: string;
  method?: string;
  extraHeaders?: Record<string, string>;
  body?: string | ArrayBuffer;
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
    const errPayload = res.json as GraphErrorPayload | undefined;
    const graphErr = errPayload?.error;
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
    const errPayload = res.json as GraphErrorPayload | undefined;
    const graphErr = errPayload?.error;
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

    const isSharePointScopeRetryable =
      res.status === 403 &&
      (opts.url.includes("/sites/") || opts.url.includes("/shares/"));

    if ((res.status === 401 || isSharePointScopeRetryable) && !opts.omitAuth) {
      const statusText = res.status === 401 ? "401 Unauthorized" : "403 Forbidden";
      this.logger.warn(
        `${method} ${safeUrl} returned ${statusText} — forcing OAuth2 token refresh and retrying once`,
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
        const ownFolderJson = ownFolderRes.json as GraphListResponse<GraphDriveItemNode> | undefined;
        for (const item of ownFolderJson?.value ?? []) {
          if (item.folder && item.name && item.parentReference?.driveId) {
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
      const sharedJson = sharedRes.json as GraphListResponse<GraphDriveItemNode> | undefined;
      for (const item of sharedJson?.value ?? []) {
        const remote = item.remoteItem ?? item;
        if (remote.folder && remote.name && remote.parentReference?.driveId) {
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
        const sitesJson = sitesRes.json as GraphListResponse<GraphSiteNode> | undefined;
        for (const site of sitesJson?.value ?? []) {
          const driveRes = await this.graphRequest({
            url: `${base}/sites/${site.id}/drive`,
            throwOnError: false,
          });
          const driveJson = driveRes.json as GraphDriveNode | undefined;
          if (driveRes.status === 200 && driveJson) {
            addVault({
              name: `${site.displayName} (SharePoint)`,
              driveId: driveJson.id,
              rootItemId: "root",
              webUrl: driveJson.webUrl ?? "",
            });

            const childrenRes = await this.graphRequest({
              url: `${base}/drives/${driveJson.id}/root/children`,
              throwOnError: false,
            });
            if (childrenRes.status === 200) {
              const childrenJson = childrenRes.json as
                | GraphListResponse<GraphDriveItemNode>
                | undefined;
              for (const child of childrenJson?.value ?? []) {
                if (child.folder && child.name) {
                  addVault({
                    name: `${child.name} (SharePoint: ${site.displayName})`,
                    driveId: driveJson.id,
                    rootItemId: child.id,
                    webUrl: child.webUrl ?? "",
                  });
                }
              }
            }
          }
        }
      }
    }

    return vaults;
  }

  /**
   * Create or resolve a vault folder in the user's OneDrive/SharePoint drive by path
   * (e.g., "ObsidianVaults/EngineeringWiki" or "Documents/Documents/EngineeringWiki"),
   * or resolve a full SharePoint / Teams / OneDrive URL if an http(s) URL is passed.
   */
  async createOrGetVaultFolder(
    vaultName: string,
    parentFolder = "ObsidianVaults"
  ): Promise<RemoteVaultTarget> {
    if (/^https?:\/\//i.test(vaultName.trim())) {
      return this.resolveVaultFromUrl(vaultName.trim());
    }
    if (/^https?:\/\//i.test(parentFolder.trim())) {
      const baseTarget = await this.resolveVaultFromUrl(parentFolder.trim());
      const cleanSub = vaultName.replace(/^\/+|\/+$/g, "");
      if (!cleanSub) return baseTarget;
      const suffixMatch = baseTarget.name.match(/\(([^)]+)\)\s*$/);
      const labelSuffix = suffixMatch ? suffixMatch[1] : "SharePoint";
      return this.ensureFolderInDrive(
        baseTarget.driveId,
        cleanSub.split("/").filter(Boolean),
        baseTarget.rootItemId,
        labelSuffix
      );
    }

    const base = this.baseUrl();
    const cleanParent = parentFolder.replace(/^\/+|\/+$/g, "");
    const cleanName = vaultName.replace(/^\/+|\/+$/g, "");
    const fullPath = cleanParent ? `${cleanParent}/${cleanName}` : cleanName;

    // Ensure each segment exists on /me/drive
    const segments = fullPath.split("/").filter(Boolean);
    let currentParentId = "root";
    let driveId = "";
    let lastItem: GraphDriveItemNode | null = null;

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
          "@microsoft.graph.conflictBehavior": "fail",
        }),
        throwOnError: false,
      });

      const createdItem = res.json as GraphDriveItemNode | undefined;
      if (res.status >= 200 && res.status < 300 && createdItem) {
        lastItem = createdItem;
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
        const fetchedItem = getRes.json as GraphDriveItemNode;
        lastItem = fetchedItem;
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
   * Resolve a SharePoint, Microsoft Teams, or OneDrive folder URL into a RemoteVaultTarget.
   * Supports:
   *  - SharePoint/Teams Forms/AllItems.aspx?id=/sites/{site}/{library}/{folder}&viewid=...
   *  - Direct SharePoint/Teams folder or library URLs (https://{tenant}.sharepoint.com/sites/{site}/{library}/{folder})
   *  - SharePoint/OneDrive sharing links (/:f:/r/..., /:f:/s/..., 1drv.ms)
   *  - Consumer OneDrive web URLs (https://onedrive.live.com/my?id=/personal/{cid}/Documents/{folder})
   */
  async resolveVaultFromUrl(rawUrl: string): Promise<RemoteVaultTarget> {
    const trimmed = rawUrl.trim();
    const parsed = new URL(trimmed);
    const base = this.baseUrl();

    const idParam =
      parsed.searchParams.get("id") || parsed.searchParams.get("RootFolder");

    const shareCandidates: string[] = [];
    if (idParam && idParam.startsWith("/")) {
      const encodedIdPath = idParam
        .split("/")
        .map((seg) => (seg ? encodeURIComponent(safeDecodeURIComponent(seg)) : ""))
        .join("/");
      shareCandidates.push(`${parsed.origin}${encodedIdPath}`);
    }

    if (/\/Forms\/[^/]+\.aspx$/i.test(parsed.pathname)) {
      const libraryRootPath = parsed.pathname.replace(/\/Forms\/[^/]+\.aspx$/i, "");
      if (libraryRootPath) {
        shareCandidates.push(`${parsed.origin}${libraryRootPath}`);
      }
    } else {
      shareCandidates.push(trimmed);
      const cleanOriginPath = `${parsed.origin}${parsed.pathname.replace(/\/+$/, "")}`;
      if (!shareCandidates.includes(cleanOriginPath)) {
        shareCandidates.push(cleanOriginPath);
      }
    }

    // 1. Try resolving directly via Microsoft Graph /shares/{u!base64url}/driveItem
    for (const candidate of shareCandidates) {
      const shareToken = encodeGraphShareToken(candidate);
      const shareRes = await this.graphRequest({
        url: `${base}/shares/${shareToken}/driveItem`,
        throwOnError: false,
      });
      if (shareRes.status === 200 && shareRes.json) {
        const rawItem = shareRes.json as GraphDriveItemNode;
        const item = rawItem.remoteItem ?? rawItem;
        const driveId = item.parentReference?.driveId;
        if (driveId && item.id) {
          const siteMatch = (idParam || parsed.pathname).match(
            /\/(?:sites|teams)\/([^/]+)/i
          );
          const siteLabel = siteMatch
            ? `SharePoint: ${safeDecodeURIComponent(siteMatch[1])}`
            : parsed.hostname.includes("sharepoint.com")
              ? "SharePoint"
              : "OneDrive";
          const isDriveRoot = item.name === "root";
          const displayName =
            (isDriveRoot && siteMatch
              ? safeDecodeURIComponent(siteMatch[1])
              : item.name) || "Vault";
          return {
            name: `${displayName} (${siteLabel})`,
            driveId,
            rootItemId: isDriveRoot ? "root" : item.id,
            webUrl: item.webUrl ?? candidate,
          };
        }
      }
    }

    // 2. SharePoint / Teams site + Document Library + subfolder resolution
    const rawServerPath =
      idParam || parsed.pathname.replace(/\/Forms\/[^/]+\.aspx$/i, "");
    let decodedSegments = rawServerPath
      .split("/")
      .filter(Boolean)
      .map(safeDecodeURIComponent);

    // Strip SharePoint sharing link prefixes such as /:f:/r/sites/... or /:f:/s/...
    if (decodedSegments[0]?.startsWith(":") && decodedSegments[0]?.endsWith(":")) {
      decodedSegments = decodedSegments.slice(1);
      if (decodedSegments[0] === "r" || decodedSegments[0] === "s") {
        decodedSegments = decodedSegments.slice(1);
      }
    }

    if (
      (decodedSegments[0] === "sites" || decodedSegments[0] === "teams") &&
      decodedSegments.length >= 2
    ) {
      const siteCollection = decodedSegments[0];
      const siteName = decodedSegments[1];
      const librarySegment = decodedSegments[2];
      let folderSegments = decodedSegments.slice(3);

      const siteRes = await this.graphRequest({
        url: `${base}/sites/${parsed.hostname}:/${siteCollection}/${encodeURIComponent(siteName)}`,
        throwOnError: true,
      });
      const site = siteRes.json as GraphSiteNode;

      const drivesRes = await this.graphRequest({
        url: `${base}/sites/${site.id}/drives`,
        throwOnError: false,
      });
      const drivesList =
        (drivesRes.json as GraphListResponse<GraphDriveNode> | undefined)?.value ?? [];

      let matchedDrive: GraphDriveNode | undefined;
      if (librarySegment && drivesList.length > 0) {
        const expectedSuffix = `/${siteCollection}/${siteName}/${librarySegment}`.toLowerCase();
        matchedDrive = drivesList.find((d) => {
          if (d.name && d.name.toLowerCase() === librarySegment.toLowerCase()) {
            return true;
          }
          if (d.webUrl) {
            try {
              const dPath = safeDecodeURIComponent(new URL(d.webUrl).pathname)
                .replace(/\/+$/, "")
                .toLowerCase();
              return dPath.endsWith(expectedSuffix);
            } catch {
              return false;
            }
          }
          return false;
        });
      }

      if (!matchedDrive) {
        const defaultDriveRes = await this.graphRequest({
          url: `${base}/sites/${site.id}/drive`,
          throwOnError: true,
        });
        matchedDrive = defaultDriveRes.json as GraphDriveNode;
        if (librarySegment) {
          const defaultWebPath = matchedDrive.webUrl
            ? safeDecodeURIComponent(new URL(matchedDrive.webUrl).pathname)
                .replace(/\/+$/, "")
                .toLowerCase()
            : "";
          const expectedSuffix = `/${siteCollection}/${siteName}/${librarySegment}`.toLowerCase();
          if (
            !defaultWebPath.endsWith(expectedSuffix) &&
            matchedDrive.name?.toLowerCase() !== librarySegment.toLowerCase()
          ) {
            folderSegments = [librarySegment, ...folderSegments];
          }
        }
      }

      const siteLabel = `SharePoint: ${site.displayName || siteName}`;
      if (folderSegments.length === 0) {
        return {
          name: `${site.displayName || siteName} (SharePoint)`,
          driveId: matchedDrive.id,
          rootItemId: "root",
          webUrl: matchedDrive.webUrl ?? trimmed,
        };
      }

      return this.ensureFolderInDrive(
        matchedDrive.id,
        folderSegments,
        "root",
        siteLabel
      );
    }

    // 3. Consumer OneDrive or OneDrive for Business /personal/{id}/Documents/{path}
    if (
      decodedSegments[0] === "personal" &&
      decodedSegments.length >= 4 &&
      decodedSegments[2].toLowerCase() === "documents"
    ) {
      const relSegments = decodedSegments.slice(3);
      const vaultFolder = relSegments[relSegments.length - 1];
      const parentPath = relSegments.slice(0, -1).join("/");
      return this.createOrGetVaultFolder(vaultFolder, parentPath);
    }

    throw new Error(
      `Could not resolve SharePoint, Teams, or OneDrive folder from URL: ${trimmed}`
    );
  }

  private async ensureFolderInDrive(
    driveId: string,
    folderSegments: string[],
    startParentId = "root",
    labelSuffix = "SharePoint"
  ): Promise<RemoteVaultTarget> {
    const base = this.baseUrl();
    const encodedRel = folderSegments.map((s) => encodeURIComponent(s)).join("/");

    if (startParentId === "root") {
      const existingRes = await this.graphRequest({
        url: `${base}/drives/${driveId}/root:/${encodedRel}`,
        throwOnError: false,
      });
      if (existingRes.status === 200 && existingRes.json) {
        const existingItem = existingRes.json as GraphDriveItemNode;
        const existingName = existingItem.name || folderSegments[folderSegments.length - 1];
        return {
          name: `${existingName} (${labelSuffix})`,
          driveId,
          rootItemId: existingItem.id,
          webUrl: existingItem.webUrl ?? "",
        };
      }
    }

    let currentParentId = startParentId;
    let lastItem: GraphDriveItemNode | null = null;
    for (let i = 0; i < folderSegments.length; i++) {
      const segment = folderSegments[i];
      const childrenUrl =
        currentParentId === "root"
          ? `${base}/drives/${driveId}/root/children`
          : `${base}/drives/${driveId}/items/${currentParentId}/children`;

      const createRes = await this.graphRequest({
        url: childrenUrl,
        method: "POST",
        extraHeaders: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: segment,
          folder: {},
        }),
        throwOnError: false,
      });

      if (createRes.status >= 200 && createRes.status < 300 && createRes.json) {
        lastItem = createRes.json as GraphDriveItemNode;
        currentParentId = lastItem.id;
      } else if (startParentId === "root") {
        const subPath = folderSegments
          .slice(0, i + 1)
          .map((s) => encodeURIComponent(s))
          .join("/");
        const getRes = await this.graphRequest({
          url: `${base}/drives/${driveId}/root:/${subPath}`,
          throwOnError: true,
        });
        lastItem = getRes.json as GraphDriveItemNode;
        currentParentId = lastItem.id;
      } else {
        const listRes = await this.graphRequest({
          url: childrenUrl,
          throwOnError: false,
        });
        const children =
          (listRes.json as GraphListResponse<GraphDriveItemNode> | undefined)?.value ?? [];
        const existingChild = children.find(
          (c) => c.folder && c.name?.toLowerCase() === segment.toLowerCase()
        );
        if (existingChild) {
          lastItem = existingChild;
          currentParentId = existingChild.id;
        } else {
          throw new Error(
            `Failed to resolve or create subfolder '${segment}' in drive ${driveId} (${createRes.status})`
          );
        }
      }
    }

    const finalName = lastItem?.name || folderSegments[folderSegments.length - 1];
    return {
      name: `${finalName} (${labelSuffix})`,
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
    const permJson = res.json as GraphListResponse<VaultPermission> | undefined;
    return permJson?.value ?? [];
  }

  async listPermissions(target: RemoteVaultTarget): Promise<VaultPermission[]> {
    const base = this.baseUrl();
    const res = await this.graphRequest({
      url: `${base}/drives/${target.driveId}/items/${target.rootItemId}/permissions`,
      throwOnError: false,
    });
    if (res.status !== 200) return [];
    const permJson = res.json as GraphListResponse<VaultPermission> | undefined;
    return permJson?.value ?? [];
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

      const deltaJson = res.json as GraphListResponse<DeltaDriveItem> | undefined;
      items.push(...(deltaJson?.value ?? []));
      if (deltaJson?.["@odata.nextLink"]) {
        url = deltaJson["@odata.nextLink"];
      } else {
        nextDeltaLink = deltaJson?.["@odata.deltaLink"] ?? "";
        break;
      }
    }

    return { items, nextDeltaLink };
  }

  private async fetchItemContentResponse(
    driveId: string,
    itemId: string
  ): Promise<{ metaRes: RequestUrlResponse; contentRes: RequestUrlResponse }> {
    const base = this.baseUrl();
    const metaRes = await this.graphRequest({
      url: `${base}/drives/${driveId}/items/${itemId}`,
    });
    const metaItem = metaRes.json as GraphDriveItemNode | undefined;
    const downloadUrl: string =
      metaItem?.["@microsoft.graph.downloadUrl"] ||
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

    return { metaRes, contentRes };
  }

  async downloadFileText(
    driveId: string,
    itemId: string
  ): Promise<{ text: string; eTag: string; lastModifiedBy?: string }> {
    const { metaRes, contentRes } = await this.fetchItemContentResponse(driveId, itemId);
    const metaItem = metaRes.json as GraphDriveItemNode | undefined;
    const author =
      metaItem?.lastModifiedBy?.user?.displayName ||
      metaItem?.lastModifiedBy?.user?.email;

    return {
      text: contentRes.text,
      eTag: metaItem?.eTag ?? "",
      lastModifiedBy: author,
    };
  }

  async downloadFileBinary(
    driveId: string,
    itemId: string
  ): Promise<{ data: ArrayBuffer; eTag: string; lastModifiedBy?: string }> {
    const { metaRes, contentRes } = await this.fetchItemContentResponse(driveId, itemId);
    const metaItem = metaRes.json as GraphDriveItemNode | undefined;
    const author =
      metaItem?.lastModifiedBy?.user?.displayName ||
      metaItem?.lastModifiedBy?.user?.email;

    return {
      data: contentRes.arrayBuffer,
      eTag: metaItem?.eTag ?? "",
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
      item: res.status >= 200 && res.status < 300 ? (res.json as DeltaDriveItem) : undefined,
    };
  }

  /**
   * Upload binary attachment content with optimistic concurrency (If-Match: eTag).
   */
  async uploadFileBinary(
    target: RemoteVaultTarget,
    relativePath: string,
    data: ArrayBuffer,
    expectedETag?: string
  ): Promise<{ status: number; item?: DeltaDriveItem }> {
    const extra: Record<string, string> = { "Content-Type": "application/octet-stream" };
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
      body: data,
      throwOnError: false,
    });

    if (res.status !== 412 && (res.status < 200 || res.status >= 300)) {
      throw new Error(
        `Graph PUT ${this.sanitizeUrl(url)} failed (${res.status}): ${this.formatErrorSummary(res)}`
      );
    }

    return {
      status: res.status,
      item: res.status >= 200 && res.status < 300 ? (res.json as DeltaDriveItem) : undefined,
    };
  }

  /**
   * Look up a single file or folder by relative path inside the vault root.
   * Returns null if the item does not exist (404).
   */
  async getItemByPath(
    target: RemoteVaultTarget,
    relativePath: string
  ): Promise<DeltaDriveItem | null> {
    const base = this.baseUrl();
    const encodedPath = relativePath
      .split("/")
      .map((seg) => encodeURIComponent(seg))
      .join("/");
    const url =
      target.rootItemId === "root"
        ? `${base}/drives/${target.driveId}/root:/${encodedPath}`
        : `${base}/drives/${target.driveId}/items/${target.rootItemId}:/${encodedPath}`;
    const res = await this.graphRequest({
      url,
      throwOnError: false,
    });
    if (res.status === 404) {
      return null;
    }
    if (res.status < 200 || res.status >= 300) {
      return null;
    }
    return (res.json as DeltaDriveItem) ?? null;
  }

  /**
   * Delete a file or folder by itemId on SharePoint / OneDrive.
   * Treats 404 (already deleted) as success.
   */
  async deleteItem(driveId: string, itemId: string): Promise<boolean> {
    const res = await this.deleteItemWithETag(driveId, itemId);
    return res.deleted;
  }

  /**
   * Delete an item by itemId with optional optimistic concurrency (If-Match: eTag).
   * Returns { deleted: false, conflict: true } on HTTP 412 when remote item changed.
   */
  async deleteItemWithETag(
    driveId: string,
    itemId: string,
    expectedETag?: string
  ): Promise<{ deleted: boolean; conflict: boolean }> {
    const base = this.baseUrl();
    const url = `${base}/drives/${driveId}/items/${itemId}`;
    const extra: Record<string, string> = {};
    if (expectedETag) {
      extra["If-Match"] = expectedETag;
    }
    const res = await this.graphRequest({
      url,
      method: "DELETE",
      extraHeaders: extra,
      throwOnError: false,
    });
    if (res.status === 412) {
      return { deleted: false, conflict: true };
    }
    if (res.status === 404) {
      return { deleted: false, conflict: false };
    }
    if (res.status < 200 || res.status >= 300) {
      throw new Error(
        `Graph DELETE ${this.sanitizeUrl(url)} failed (${res.status}): ${this.formatErrorSummary(res)}`
      );
    }
    return { deleted: true, conflict: false };
  }

  /**
   * Delete a file or folder by relative path inside the vault root on SharePoint / OneDrive.
   * Useful for deleting folders whose itemId wasn't cached.
   */
  async deleteItemByPath(
    target: RemoteVaultTarget,
    relativePath: string
  ): Promise<boolean> {
    const base = this.baseUrl();
    const encodedPath = relativePath
      .split("/")
      .map((seg) => encodeURIComponent(seg))
      .join("/");
    const url = `${base}/drives/${target.driveId}/items/${target.rootItemId}:/${encodedPath}`;
    const res = await this.graphRequest({
      url,
      method: "DELETE",
      throwOnError: false,
    });
    if (res.status === 404) {
      return false;
    }
    if (res.status < 200 || res.status >= 300) {
      throw new Error(
        `Graph DELETE ${this.sanitizeUrl(url)} failed (${res.status}): ${this.formatErrorSummary(res)}`
      );
    }
    return true;
  }

  /**
   * List version history for a note (author + timestamp)
   */
  async listVersions(driveId: string, itemId: string): Promise<DriveItemVersion[]> {
    const base = this.baseUrl();
    const res = await this.graphRequest({
      url: `${base}/drives/${driveId}/items/${itemId}/versions`,
    });
    const versionJson = res.json as GraphListResponse<DriveItemVersion> | undefined;
    return versionJson?.value ?? [];
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
