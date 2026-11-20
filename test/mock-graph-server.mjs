import http from "node:http";
import crypto from "node:crypto";

/**
 * Lightweight Microsoft Graph (/v1.0/drives & /v1.0/me) SharePoint/OneDrive emulator for E2E testing.
 * Validates real Entra ID Bearer access tokens against live https://graph.microsoft.com/v1.0/me
 * (with in-memory caching) so author attribution and permissions use real Entra ID identities.
 */
export function createMockGraphServer({ port = 0, host = "0.0.0.0", baseUrlOverride = null } = {}) {
  const tokenCache = new Map();
  // drives: Map<driveId, { id, ownerEmail, ownerName, items: Map<itemId, Item>, deltaSeq: Map<itemId, seq> }>
  const drives = new Map();
  const followedSites = new Map();
  const registeredShares = new Map();
  const forceShareFailureHosts = new Set();
  const authCodes = new Map();
  const refreshTokens = new Map();
  let globalSeq = 1;

  function makeMockJwt(user) {
    const header = Buffer.from(JSON.stringify({ alg: "none", typ: "JWT" })).toString("base64url");
    const payload = Buffer.from(
      JSON.stringify({
        oid: user.id,
        name: user.displayName,
        upn: user.email,
        nonce: crypto.randomUUID(),
      })
    ).toString("base64url");
    return `${header}.${payload}.`;
  }

  function ensureSharePointSite({
    slug,
    displayName,
    hostname = "contoso.sharepoint.com",
    kind = "sites",
    library = "Shared Documents",
    folders = [],
  }) {
    const cleanSlug = slug.toLowerCase();
    const siteId = `site-${cleanSlug}`;
    const spDriveId = `sp-drive-${cleanSlug}`;
    const webUrl = `https://${hostname}/${kind}/${slug}`;
    const libraryWebUrl = `${webUrl}/${encodeURIComponent(library).replace(/%20/g, "%20")}`;

    if (!drives.has(spDriveId)) {
      const items = new Map();
      items.set("root", {
        id: "root",
        name: "root",
        folder: { childCount: 0 },
        parentReference: { driveId: spDriveId, id: "", path: "" },
        webUrl: libraryWebUrl,
        permissions: [],
      });
      drives.set(spDriveId, {
        id: spDriveId,
        ownerEmail: `admin@${hostname}`,
        ownerName: displayName || slug,
        items,
        deltaSeq: new Map(),
      });
    }

    const spDrive = drives.get(spDriveId);
    const createdFolders = [];
    for (const folderRel of folders) {
      const folder = ensureFolderHierarchy(
        spDrive,
        "root",
        folderRel.split("/").filter(Boolean)
      );
      createdFolders.push(folder);
    }

    const siteRecord = {
      id: siteId,
      slug,
      kind,
      hostname,
      displayName: displayName || slug,
      webUrl,
      libraryWebUrl,
      driveId: spDriveId,
    };
    followedSites.set(siteId, siteRecord);
    return { ...siteRecord, createdFolders };
  }

  function ensureUserDrive(user) {
    const driveId = `drive-${user.id}`;
    if (!drives.has(driveId)) {
      const rootItem = {
        id: "root",
        name: "root",
        folder: { childCount: 0 },
        parentReference: { driveId, id: "", path: "" },
        webUrl: `https://sharepoint.local/drives/${driveId}/root`,
        permissions: [
          {
            id: "perm-owner",
            roles: ["owner", "write"],
            grantedToV2: {
              user: { displayName: user.displayName, email: user.email },
            },
          },
        ],
      };
      const items = new Map();
      items.set("root", rootItem);
      drives.set(driveId, {
        id: driveId,
        ownerEmail: user.email,
        ownerName: user.displayName,
        items,
        deltaSeq: new Map(), // itemId -> seq
      });
    }
    return drives.get(driveId);
  }

  async function authenticateRequest(req) {
    const auth = req.headers["authorization"] || "";
    const match = auth.match(/^Bearer\s+(.+)$/i);
    if (!match) {
      return null;
    }
    const token = match[1].trim();
    const tokenHash = crypto.createHash("sha256").update(token).digest("hex");
    if (tokenCache.has(tokenHash)) {
      return tokenCache.get(tokenHash);
    }

    // Fast offline decode for unsigned mock JWTs (alg: "none")
    try {
      const parts = token.split(".");
      if (parts.length >= 2) {
        const hdr = JSON.parse(Buffer.from(parts[0], "base64url").toString("utf8"));
        if (hdr.alg === "none") {
          const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));
          const user = {
            id: payload.oid || payload.sub || "unknown",
            displayName: payload.name || payload.upn || "User",
            email: (payload.upn || payload.preferred_username || payload.email || "").toLowerCase(),
            upn: payload.upn || payload.preferred_username,
          };
          tokenCache.set(tokenHash, user);
          return user;
        }
      }
    } catch {
      // Fall through to live Graph / standard JWT decode
    }

    // Verify against real Microsoft Graph /v1.0/me first
    try {
      const res = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (res.ok) {
        const me = await res.json();
        const user = {
          id: me.id,
          displayName: me.displayName || me.userPrincipalName,
          email: (me.mail || me.userPrincipalName || "").toLowerCase(),
          upn: me.userPrincipalName,
        };
        tokenCache.set(tokenHash, user);
        return user;
      }
    } catch {
      // Fallback to JWT claim decoding if offline
    }

    try {
      const payload = JSON.parse(
        Buffer.from(token.split(".")[1], "base64url").toString("utf8")
      );
      const user = {
        id: payload.oid || payload.sub || "unknown",
        displayName: payload.name || payload.upn || "User",
        email: (payload.upn || payload.preferred_username || payload.email || "").toLowerCase(),
        upn: payload.upn || payload.preferred_username,
      };
      tokenCache.set(tokenHash, user);
      return user;
    } catch {
      return null;
    }
  }

  function findFolderByPath(drive, parentId, relPath) {
    const segments = relPath.split("/").filter(Boolean).map(decodeURIComponent);
    let current = drive.items.get(parentId);
    if (!current) return null;
    for (const seg of segments) {
      let next = null;
      for (const item of drive.items.values()) {
        if (item.parentReference?.id === current.id && item.name === seg && !item.deleted) {
          next = item;
          break;
        }
      }
      if (!next) return null;
      current = next;
    }
    return current;
  }

  function ensureFolderHierarchy(drive, parentId, segments) {
    let current = drive.items.get(parentId);
    for (const seg of segments) {
      let next = null;
      for (const item of drive.items.values()) {
        if (
          item.parentReference?.id === current.id &&
          item.name === seg &&
          item.folder &&
          !item.deleted
        ) {
          next = item;
          break;
        }
      }
      if (!next) {
        const id = `folder-${crypto.randomUUID().slice(0, 8)}`;
        const parentPath =
          current.id === "root"
            ? `/drives/${drive.id}/root:`
            : `${current.parentReference.path}/${current.name}`;
        next = {
          id,
          name: seg,
          eTag: `"etag-${globalSeq++}"`,
          folder: { childCount: 0 },
          parentReference: {
            driveId: drive.id,
            id: current.id,
            path: parentPath,
          },
          webUrl: `https://sharepoint.local/drives/${drive.id}/items/${id}`,
          permissions: [
            {
              id: "perm-owner",
              roles: ["owner", "write"],
              grantedToV2: {
                user: { displayName: drive.ownerName, email: drive.ownerEmail },
              },
            },
          ],
        };
        drive.items.set(id, next);
        drive.deltaSeq.set(id, globalSeq++);
      }
      current = next;
    }
    return current;
  }

  function isDescendantOf(drive, item, ancestorId) {
    if (ancestorId === "root") return item.id !== "root";
    let cur = item;
    while (cur && cur.parentReference?.id) {
      if (cur.parentReference.id === ancestorId) return true;
      cur = drive.items.get(cur.parentReference.id);
    }
    return false;
  }

  function sendJson(res, status, obj) {
    const body = JSON.stringify(obj);
    res.writeHead(status, {
      "Content-Type": "application/json; charset=utf-8",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "*",
      "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
    });
    res.end(body);
  }

  function sendBuffer(res, status, buf, contentType = "application/octet-stream") {
    res.writeHead(status, {
      "Content-Type": contentType,
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "*",
      "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
    });
    res.end(buf);
  }

  function markDeletedRecursive(drive, targetItem) {
    targetItem.deleted = { state: "deleted" };
    targetItem.eTag = `"etag-${globalSeq++}"`;
    drive.deltaSeq.set(targetItem.id, globalSeq++);
    if (targetItem.folder) {
      for (const child of drive.items.values()) {
        if (child.id !== targetItem.id && isDescendantOf(drive, child, targetItem.id)) {
          child.deleted = { state: "deleted" };
          child.eTag = `"etag-${globalSeq++}"`;
          drive.deltaSeq.set(child.id, globalSeq++);
        }
      }
    }
  }

  const server = http.createServer(async (req, res) => {
    if (req.method === "OPTIONS") {
      res.writeHead(204, {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
      });
      res.end();
      return;
    }

    const reqUrl = new URL(req.url, `http://${req.headers.host || "localhost"}`);
    const pathname = reqUrl.pathname;

    if (pathname === "/health") {
      sendJson(res, 200, { ok: true });
      return;
    }

    // Unauthenticated OAuth 2.0 PKCE token endpoint: POST /:tenantId/oauth2/v2.0/token
    const oauthTokenMatch = pathname.match(/^\/([^/]+)\/oauth2\/v2\.0\/token$/);
    if (req.method === "POST" && oauthTokenMatch) {
      const tenantId = decodeURIComponent(oauthTokenMatch[1]);
      const chunks = [];
      for await (const chunk of req) {
        chunks.push(chunk);
      }
      const formBody = Buffer.concat(chunks).toString("utf8");
      const params = new URLSearchParams(formBody);
      const grantType = params.get("grant_type");
      const clientId = params.get("client_id") || "";

      if (grantType === "authorization_code") {
        const code = params.get("code") || "";
        const codeVerifier = params.get("code_verifier") || "";
        const redirectUri = params.get("redirect_uri") || "";
        const record = authCodes.get(code);
        if (!record) {
          sendJson(res, 400, { error: "invalid_grant", error_description: "Unknown or expired authorization code" });
          return;
        }
        authCodes.delete(code);

        if (record.tenantId !== tenantId || record.clientId !== clientId || record.redirectUri !== redirectUri) {
          sendJson(res, 400, { error: "invalid_client", error_description: "Tenant, client_id, or redirect_uri mismatch" });
          return;
        }

        const computedChallenge = crypto
          .createHash("sha256")
          .update(codeVerifier)
          .digest("base64url");
        if (computedChallenge !== record.codeChallenge) {
          sendJson(res, 400, { error: "invalid_grant", error_description: "PKCE code_verifier mismatch" });
          return;
        }

        const accessToken = makeMockJwt(record.user);
        const newRefreshToken = `refresh-${crypto.randomUUID()}`;
        refreshTokens.set(newRefreshToken, {
          tenantId,
          clientId,
          scope: record.scope,
          user: record.user,
        });
        ensureUserDrive(record.user);

        sendJson(res, 200, {
          token_type: "Bearer",
          scope: record.scope,
          expires_in: 3600,
          access_token: accessToken,
          refresh_token: newRefreshToken,
        });
        return;
      }

      if (grantType === "refresh_token") {
        const incomingRefresh = params.get("refresh_token") || "";
        const record = refreshTokens.get(incomingRefresh);
        if (!record) {
          sendJson(res, 400, { error: "invalid_grant", error_description: "Invalid refresh token" });
          return;
        }
        refreshTokens.delete(incomingRefresh);
        const rotatedRefresh = `refresh-${crypto.randomUUID()}`;
        refreshTokens.set(rotatedRefresh, record);
        sendJson(res, 200, {
          token_type: "Bearer",
          scope: record.scope,
          expires_in: 3600,
          access_token: makeMockJwt(record.user),
          refresh_token: rotatedRefresh,
        });
        return;
      }

      sendJson(res, 400, { error: "unsupported_grant_type" });
      return;
    }

    const user = await authenticateRequest(req);
    if (!user) {
      sendJson(res, 401, { error: { code: "InvalidAuthenticationToken", message: "Unauthorized" } });
      return;
    }

    const userDrive = ensureUserDrive(user);
    const chunks = [];
    for await (const chunk of req) {
      chunks.push(chunk);
    }
    const rawBuffer = Buffer.concat(chunks);
    const rawBody = rawBuffer.toString("utf8");

    const effectiveBase =
      baseUrlOverride || `http://${req.headers.host || "localhost"}/v1.0`;

    // 1. GET /v1.0/me
    if (req.method === "GET" && pathname === "/v1.0/me") {
      sendJson(res, 200, {
        id: user.id,
        displayName: user.displayName,
        mail: user.email,
        userPrincipalName: user.upn || user.email,
      });
      return;
    }

    // 2. POST /v1.0/me/drive/root/children or /v1.0/me/drive/items/:parentId/children
    const createChildMatch = pathname.match(
      /^\/v1\.0\/me\/drive\/(?:(root)|items\/([^/]+))\/children$/
    );
    if (req.method === "POST" && createChildMatch) {
      const parentId = createChildMatch[1] ? "root" : createChildMatch[2];
      const payload = JSON.parse(rawBody || "{}");
      const folder = ensureFolderHierarchy(userDrive, parentId, [payload.name]);
      sendJson(res, 201, folder);
      return;
    }

    // 3. GET /v1.0/me/drive/root:/:folderPath:/children
    const ownChildrenMatch = pathname.match(/^\/v1\.0\/me\/drive\/root:\/(.+?):\/children$/);
    if (req.method === "GET" && ownChildrenMatch) {
      const folderPath = ownChildrenMatch[1];
      const folder = findFolderByPath(userDrive, "root", folderPath);
      if (!folder) {
        sendJson(res, 200, { value: [] });
        return;
      }
      const children = [];
      for (const item of userDrive.items.values()) {
        if (item.parentReference?.id === folder.id && !item.deleted) {
          children.push(item);
        }
      }
      sendJson(res, 200, { value: children });
      return;
    }

    // 4. GET /v1.0/me/drive/sharedWithMe
    if (req.method === "GET" && pathname === "/v1.0/me/drive/sharedWithMe") {
      const shared = [];
      for (const drive of drives.values()) {
        for (const item of drive.items.values()) {
          if (item.folder && item.permissions) {
            const hasPerm = item.permissions.some(
              (p) =>
                drive.ownerEmail !== user.email &&
                (p.grantedToV2?.user?.email?.toLowerCase() === user.email ||
                  p.invitation?.email?.toLowerCase() === user.email)
            );
            if (hasPerm) {
              shared.push({
                id: item.id,
                name: item.name,
                folder: item.folder,
                parentReference: item.parentReference,
                webUrl: item.webUrl,
                remoteItem: {
                  id: item.id,
                  name: item.name,
                  folder: item.folder,
                  parentReference: item.parentReference,
                  webUrl: item.webUrl,
                },
              });
            }
          }
        }
      }
      sendJson(res, 200, { value: shared });
      return;
    }

    // 5. GET /v1.0/me/followedSites
    if (req.method === "GET" && pathname === "/v1.0/me/followedSites") {
      const list = Array.from(followedSites.values()).map((s) => ({
        id: s.id,
        displayName: s.displayName,
        webUrl: s.webUrl,
      }));
      sendJson(res, 200, { value: list });
      return;
    }

    // 5a. GET /v1.0/sites/:hostname:/:sitePath (fallback SharePoint site lookup)
    const siteByPathMatch = pathname.match(/^\/v1\.0\/sites\/([^/:]+):\/(.+)$/);
    if (req.method === "GET" && siteByPathMatch) {
      const hostname = decodeURIComponent(siteByPathMatch[1]);
      const sitePath = decodeURIComponent(siteByPathMatch[2]).replace(/^\/+|\/+$/g, "");
      const pathParts = sitePath.split("/");
      const kind = pathParts[0] || "sites";
      const slug = pathParts[1] || pathParts[0];
      let found = Array.from(followedSites.values()).find(
        (s) => s.hostname.toLowerCase() === hostname.toLowerCase() && s.slug.toLowerCase() === slug.toLowerCase()
      );
      if (!found) {
        found = ensureSharePointSite({ slug, displayName: slug, hostname, kind });
      }
      sendJson(res, 200, {
        id: found.id,
        displayName: found.displayName,
        webUrl: found.webUrl,
      });
      return;
    }

    // 5a-2. GET /v1.0/sites/:siteId/drive
    const siteDriveMatch = pathname.match(/^\/v1\.0\/sites\/([^/:]+)\/drive$/);
    if (req.method === "GET" && siteDriveMatch) {
      const siteId = decodeURIComponent(siteDriveMatch[1]);
      const site = followedSites.get(siteId);
      if (!site) {
        sendJson(res, 404, { error: { code: "itemNotFound", message: "Site not found" } });
        return;
      }
      sendJson(res, 200, {
        id: site.driveId,
        webUrl: site.libraryWebUrl,
      });
      return;
    }

    // 5b. GET /v1.0/shares/:shareToken/driveItem (SharePoint / Teams / OneDrive URL resolution)
    const shareItemMatch = pathname.match(/^\/v1\.0\/shares\/([^/]+)\/driveItem$/);
    if (req.method === "GET" && shareItemMatch) {
      const shareToken = shareItemMatch[1];
      const b64 = shareToken.replace(/^u!/, "");
      let decodedUrl = "";
      try {
        decodedUrl = Buffer.from(b64, "base64url").toString("utf8");
      } catch {
        sendJson(res, 400, { error: { code: "invalidRequest", message: "Bad share token" } });
        return;
      }

      // Registered OneDrive / SharePoint short links (e.g. https://1drv.ms/... or https://onedrive.live.com/...)
      if (registeredShares.has(decodedUrl)) {
        const { driveId, itemId } = registeredShares.get(decodedUrl);
        const drive = drives.get(driveId);
        const item = drive?.items.get(itemId);
        if (drive && item && !item.deleted) {
          sendJson(res, 200, {
            ...item,
            parentReference: {
              ...item.parentReference,
              driveId,
            },
          });
          return;
        }
      }

      try {
        const parsedShareUrl = new URL(decodedUrl);
        if (forceShareFailureHosts.has(parsedShareUrl.hostname.toLowerCase())) {
          sendJson(res, 403, { error: { code: "accessDenied", message: "Simulated share token denial" } });
          return;
        }
      } catch {
        // ignore URL parse error
      }

      // OneDrive for Business / Personal URL: https://<host>/personal/<slug>/Documents/<folderPath>
      const personalMatch = decodedUrl.match(
        /^https?:\/\/([^/]+)\/personal\/([^/]+)\/Documents\/(.+)$/i
      );
      if (personalMatch) {
        const folderPath = decodeURIComponent(personalMatch[3]);
        const targetFolder = ensureFolderHierarchy(
          userDrive,
          "root",
          folderPath.split("/").filter(Boolean)
        );
        sendJson(res, 200, {
          ...targetFolder,
          parentReference: {
            ...targetFolder.parentReference,
            driveId: userDrive.id,
          },
          webUrl: decodedUrl,
        });
        return;
      }

      const spMatch = decodedUrl.match(
        /^https?:\/\/([^/]+)\/(sites|teams)\/([^/]+)\/([^/]+)(?:\/(.+))?$/i
      );
      if (spMatch) {
        const hostname = spMatch[1];
        const kind = spMatch[2];
        const siteSlug = decodeURIComponent(spMatch[3]);
        const library = decodeURIComponent(spMatch[4]);
        const folderPath = spMatch[5] ? decodeURIComponent(spMatch[5]) : "";
        const existingSite = Array.from(followedSites.values()).find(
          (s) => s.slug.toLowerCase() === siteSlug.toLowerCase()
        );
        if (!existingSite) {
          ensureSharePointSite({
            slug: siteSlug,
            displayName: siteSlug,
            hostname,
            kind,
            library,
          });
        }
        const spDriveId = `sp-drive-${siteSlug.toLowerCase()}`;
        const spDrive = drives.get(spDriveId);
        const targetFolder = folderPath
          ? ensureFolderHierarchy(
              spDrive,
              "root",
              folderPath.split("/").filter(Boolean)
            )
          : spDrive.items.get("root");
        sendJson(res, 200, {
          ...targetFolder,
          parentReference: {
            ...targetFolder.parentReference,
            driveId: spDriveId,
          },
          webUrl: decodedUrl,
        });
        return;
      }
      sendJson(res, 404, { error: { code: "itemNotFound", message: "Share item not found" } });
      return;
    }

    // 5c. POST /v1.0/drives/:driveId/root/children or /v1.0/drives/:driveId/items/:parentId/children
    const driveCreateChildMatch = pathname.match(
      /^\/v1\.0\/drives\/([^/]+)\/(?:(root)|items\/([^/]+))\/children$/
    );
    if (req.method === "POST" && driveCreateChildMatch) {
      const driveId = driveCreateChildMatch[1];
      const parentId = driveCreateChildMatch[2] ? "root" : driveCreateChildMatch[3];
      const drive = drives.get(driveId);
      if (!drive) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const payload = JSON.parse(rawBody || "{}");
      const conflictBehavior = payload["@microsoft.graph.conflictBehavior"] || "fail";
      if (conflictBehavior === "fail") {
        for (const item of drive.items.values()) {
          if (
            item.parentReference?.id === parentId &&
            item.name.toLowerCase() === String(payload.name || "").toLowerCase() &&
            !item.deleted
          ) {
            sendJson(res, 409, {
              error: { code: "nameAlreadyExists", message: "Name already exists" },
            });
            return;
          }
        }
      }
      const folder = ensureFolderHierarchy(drive, parentId, [payload.name]);
      sendJson(res, 201, folder);
      return;
    }

    // 5d. GET /v1.0/drives/:driveId/root/children or /v1.0/drives/:driveId/items/:parentId/children
    if (req.method === "GET" && driveCreateChildMatch) {
      const driveId = driveCreateChildMatch[1];
      const parentId = driveCreateChildMatch[2] ? "root" : driveCreateChildMatch[3];
      const drive = drives.get(driveId);
      if (!drive) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const children = [];
      for (const item of drive.items.values()) {
        if (item.parentReference?.id === parentId && item.id !== "root" && !item.deleted) {
          children.push(item);
        }
      }
      sendJson(res, 200, { value: children });
      return;
    }

    // 5e. GET /v1.0/drives/:driveId/root:/:subPath
    const driveByPathMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/root:\/(.+?)$/);
    if (req.method === "GET" && driveByPathMatch && !pathname.endsWith(":/children")) {
      const driveId = driveByPathMatch[1];
      const subPath = driveByPathMatch[2].replace(/:$/, "");
      const drive = drives.get(driveId);
      if (!drive) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const item = findFolderByPath(drive, "root", subPath);
      if (!item || item.deleted) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      sendJson(res, 200, item);
      return;
    }

    // 6. POST /v1.0/drives/:driveId/items/:itemId/invite
    const inviteMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/invite$/);
    if (req.method === "POST" && inviteMatch) {
      const [, driveId, itemId] = inviteMatch;
      const drive = drives.get(driveId);
      const item = drive?.items.get(itemId);
      if (!drive || !item) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const body = JSON.parse(rawBody || "{}");
      const roles = body.roles || ["write"];
      const recipients = body.recipients || [];
      item.permissions = item.permissions || [];
      const added = [];
      for (const r of recipients) {
        const email = (r.email || "").toLowerCase();
        if (!email) continue;
        const perm = {
          id: `perm-${crypto.randomUUID().slice(0, 8)}`,
          roles,
          grantedToV2: {
            user: {
              displayName: email.split("@")[0],
              email,
            },
          },
          invitation: { email },
        };
        item.permissions.push(perm);
        added.push(perm);
      }
      sendJson(res, 200, { value: added });
      return;
    }

    // 7. GET /v1.0/drives/:driveId/items/:itemId/permissions
    const permsMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/permissions$/);
    if (req.method === "GET" && permsMatch) {
      const [, driveId, itemId] = permsMatch;
      const drive = drives.get(driveId);
      const item = drive?.items.get(itemId);
      if (!drive || !item) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      sendJson(res, 200, { value: item.permissions || [] });
      return;
    }

    // 8. GET /v1.0/drives/:driveId/items/:rootItemId/delta
    const deltaMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/delta$/);
    if (req.method === "GET" && deltaMatch) {
      const [, driveId, rootItemId] = deltaMatch;
      const drive = drives.get(driveId);
      if (!drive) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const sinceToken = Number(reqUrl.searchParams.get("token") || "0");
      const changed = [];
      let maxSeq = sinceToken;

      for (const [id, item] of drive.items.entries()) {
        const seq = drive.deltaSeq.get(id) || 0;
        if (seq > sinceToken && isDescendantOf(drive, item, rootItemId)) {
          changed.push({
            id: item.id,
            name: item.name,
            eTag: item.eTag,
            cTag: item.cTag,
            deleted: item.deleted,
            folder: item.folder,
            file: item.file,
            parentReference: item.parentReference,
            lastModifiedDateTime: item.lastModifiedDateTime,
            lastModifiedBy: item.lastModifiedBy,
          });
          if (seq > maxSeq) maxSeq = seq;
        }
      }

      const nextDeltaLink = `${effectiveBase}/drives/${driveId}/items/${rootItemId}/delta?token=${maxSeq}`;
      sendJson(res, 200, {
        value: changed,
        "@odata.deltaLink": nextDeltaLink,
      });
      return;
    }

    // 9. PUT /v1.0/drives/:driveId/items/:rootItemId:/:relativePath:/content
    const putContentMatch = pathname.match(
      /^\/v1\.0\/drives\/([^/]+)\/items\/([^/:]+):\/(.+?):\/content$/
    );
    if (req.method === "PUT" && putContentMatch) {
      const [, driveId, rootItemId, encodedRelPath] = putContentMatch;
      const drive = drives.get(driveId);
      if (!drive) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const segments = encodedRelPath.split("/").filter(Boolean).map(decodeURIComponent);
      const fileName = segments.pop();
      const parentFolder = ensureFolderHierarchy(drive, rootItemId, segments);

      let existingFile = null;
      for (const item of drive.items.values()) {
        if (
          item.parentReference?.id === parentFolder.id &&
          item.name === fileName &&
          !item.folder
        ) {
          existingFile = item;
          break;
        }
      }

      const ifMatch = req.headers["if-match"];
      if (existingFile && !existingFile.deleted) {
        // Optimistic concurrency check:
        // If caller sent If-Match and it doesn't match current eTag, OR if caller omitted If-Match
        // on an already-existing file modified by someone else, return 412 Precondition Failed!
        if (ifMatch && ifMatch !== existingFile.eTag) {
          sendJson(res, 412, {
            error: {
              code: "notAllowed",
              message: `Precondition Failed: expected ${ifMatch}, current is ${existingFile.eTag}`,
            },
          });
          return;
        }
      }

      const nowIso = new Date().toISOString();
      const newETag = `"etag-${globalSeq++}"`;
      const authorInfo = {
        user: {
          displayName: user.displayName,
          email: user.email,
        },
      };

      if (!existingFile) {
        const id = `file-${crypto.randomUUID().slice(0, 8)}`;
        const parentPath =
          parentFolder.id === "root"
            ? `/drives/${drive.id}/root:`
            : `${parentFolder.parentReference.path}/${parentFolder.name}`;
        existingFile = {
          id,
          name: fileName,
          eTag: newETag,
          cTag: newETag,
          file: { mimeType: "text/markdown" },
          parentReference: {
            driveId: drive.id,
            id: parentFolder.id,
            path: parentPath,
          },
          content: rawBody,
          contentBuffer: rawBuffer,
          lastModifiedDateTime: nowIso,
          lastModifiedBy: authorInfo,
          versions: [
            {
              id: "1.0",
              lastModifiedDateTime: nowIso,
              lastModifiedBy: authorInfo,
              size: rawBuffer.length,
              content: rawBody,
              contentBuffer: rawBuffer,
            },
          ],
        };
        drive.items.set(id, existingFile);
        drive.deltaSeq.set(id, globalSeq++);
      } else {
        existingFile.deleted = undefined;
        existingFile.eTag = newETag;
        existingFile.cTag = newETag;
        existingFile.content = rawBody;
        existingFile.contentBuffer = rawBuffer;
        existingFile.lastModifiedDateTime = nowIso;
        existingFile.lastModifiedBy = authorInfo;
        existingFile.versions = existingFile.versions || [];
        const nextVerNum = `${existingFile.versions.length + 1}.0`;
        existingFile.versions.unshift({
          id: nextVerNum,
          lastModifiedDateTime: nowIso,
          lastModifiedBy: authorInfo,
          size: rawBuffer.length,
          content: rawBody,
          contentBuffer: rawBuffer,
        });
        drive.deltaSeq.set(existingFile.id, globalSeq++);
      }

      sendJson(res, 200, {
        id: existingFile.id,
        name: existingFile.name,
        eTag: existingFile.eTag,
        cTag: existingFile.cTag,
        file: existingFile.file,
        parentReference: existingFile.parentReference,
        lastModifiedDateTime: existingFile.lastModifiedDateTime,
        lastModifiedBy: existingFile.lastModifiedBy,
      });
      return;
    }

    // 10. GET /v1.0/drives/:driveId/items/:itemId/versions/:versionId/content
    const versionContentMatch = pathname.match(
      /^\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/versions\/([^/]+)\/content$/
    );
    if (req.method === "GET" && versionContentMatch) {
      const [, driveId, itemId, versionId] = versionContentMatch;
      const drive = drives.get(driveId);
      const item = drive?.items.get(itemId);
      const ver = item?.versions?.find((v) => v.id === versionId);
      if (!ver) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      sendBuffer(
        res,
        200,
        ver.contentBuffer || Buffer.from(ver.content || "", "utf8"),
        "text/plain; charset=utf-8"
      );
      return;
    }

    // 11. GET /v1.0/drives/:driveId/items/:itemId/versions
    const versionsMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/versions$/);
    if (req.method === "GET" && versionsMatch) {
      const [, driveId, itemId] = versionsMatch;
      const drive = drives.get(driveId);
      const item = drive?.items.get(itemId);
      if (!item) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const versions = (item.versions || []).map((v) => ({
        id: v.id,
        lastModifiedDateTime: v.lastModifiedDateTime,
        lastModifiedBy: v.lastModifiedBy,
        size: v.size,
      }));
      sendJson(res, 200, { value: versions });
      return;
    }

    // 12. GET /v1.0/drives/:driveId/items/:itemId/content
    const itemContentMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/content$/);
    if (req.method === "GET" && itemContentMatch) {
      const [, driveId, itemId] = itemContentMatch;
      const drive = drives.get(driveId);
      const item = drive?.items.get(itemId);
      if (!item || item.deleted) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      sendBuffer(
        res,
        200,
        item.contentBuffer || Buffer.from(item.content ?? "", "utf8"),
        "application/octet-stream"
      );
      return;
    }

    // 13. DELETE /v1.0/drives/:driveId/items/:rootItemId:/:relativePath
    const deleteByPathMatch = pathname.match(
      /^\/v1\.0\/drives\/([^/]+)\/items\/([^/:]+):\/(.+?)$/
    );
    if (req.method === "DELETE" && deleteByPathMatch) {
      const [, driveId, rootItemId, encodedRelPath] = deleteByPathMatch;
      const drive = drives.get(driveId);
      if (!drive) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      const targetItem = findFolderByPath(drive, rootItemId, encodedRelPath);
      if (!targetItem || targetItem.deleted) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      markDeletedRecursive(drive, targetItem);
      res.writeHead(204, {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
      });
      res.end();
      return;
    }

    // 14. DELETE /v1.0/drives/:driveId/items/:itemId
    const deleteItemMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/items\/([^/:]+)$/);
    if (req.method === "DELETE" && deleteItemMatch) {
      const [, driveId, itemId] = deleteItemMatch;
      const drive = drives.get(driveId);
      const item = drive?.items.get(itemId);
      if (!item || item.deleted) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      markDeletedRecursive(drive, item);
      res.writeHead(204, {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
      });
      res.end();
      return;
    }

    // 15. GET /v1.0/drives/:driveId/items/:itemId
    const itemMetaMatch = pathname.match(/^\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)$/);
    if (req.method === "GET" && itemMetaMatch) {
      const [, driveId, itemId] = itemMetaMatch;
      const drive = drives.get(driveId);
      const item = drive?.items.get(itemId);
      if (!item || item.deleted) {
        sendJson(res, 404, { error: { code: "itemNotFound" } });
        return;
      }
      sendJson(res, 200, {
        id: item.id,
        name: item.name,
        eTag: item.eTag,
        cTag: item.cTag,
        file: item.file,
        folder: item.folder,
        parentReference: item.parentReference,
        lastModifiedDateTime: item.lastModifiedDateTime,
        lastModifiedBy: item.lastModifiedBy,
        "@microsoft.graph.downloadUrl": `${effectiveBase}/drives/${driveId}/items/${itemId}/content`,
      });
      return;
    }

    sendJson(res, 404, { error: { code: "Request_ResourceNotFound", path: pathname } });
  });

  return {
    server,
    drives,
    issueAuthCode(authorizeUrl, user) {
      const parsed = new URL(authorizeUrl);
      const tenantId = decodeURIComponent(parsed.pathname.split("/").filter(Boolean)[0] || "common");
      const clientId = parsed.searchParams.get("client_id") || "";
      const scope = parsed.searchParams.get("scope") || "";
      const codeChallenge = parsed.searchParams.get("code_challenge") || "";
      const codeChallengeMethod = parsed.searchParams.get("code_challenge_method") || "";
      const redirectUri = parsed.searchParams.get("redirect_uri") || "";
      const state = parsed.searchParams.get("state") || "";
      const code = `pkce-code-${crypto.randomUUID()}`;
      authCodes.set(code, {
        tenantId,
        clientId,
        scope,
        codeChallenge,
        codeChallengeMethod,
        redirectUri,
        state,
        user,
      });
      return {
        code,
        state,
        tenantId,
        clientId,
        scope,
        codeChallenge,
        codeChallengeMethod,
        redirectUri,
      };
    },
    seedSharePointSite(opts) {
      return ensureSharePointSite(opts);
    },
    registerShareUrl(shareUrl, driveId, itemId) {
      registeredShares.set(shareUrl, { driveId, itemId });
    },
    setForceShareFailureHost(hostname, enabled = true) {
      if (enabled) forceShareFailureHosts.add(hostname.toLowerCase());
      else forceShareFailureHosts.delete(hostname.toLowerCase());
    },
    start() {
      return new Promise((resolve, reject) => {
        server.listen(port, host, () => {
          const addr = server.address();
          resolve(addr.port);
        });
        server.on("error", reject);
      });
    },
    stop() {
      return new Promise((resolve) => server.close(() => resolve()));
    },
  };
}
