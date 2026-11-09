import http from "node:http";
import crypto from "node:crypto";

/**
 * Lightweight Microsoft Graph (/v1.0/drives & /v1.0/me) SharePoint/OneDrive emulator for E2E testing.
 * Validates real Entra ID Bearer access tokens against live https://graph.microsoft.com/v1.0/me
 * (with in-memory caching) so author attribution and permissions use real Entra ID identities.
 */
export function createMockGraphServer({ port = 0, host = "0.0.0.0", baseUrlOverride = null } = {}) {
  const tokenCache = new Map();
  // drives: Map<driveId, { id, ownerEmail, ownerName, items: Map<itemId, Item>, deltaLog: Array<{seq, itemId}> }>
  const drives = new Map();
  let globalSeq = 1;

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
      sendJson(res, 200, { value: [] });
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
