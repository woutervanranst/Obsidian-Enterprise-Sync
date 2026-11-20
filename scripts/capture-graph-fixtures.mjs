#!/usr/bin/env node
/**
 * Interactive Microsoft Graph Anonymized Fixture Capture Tool
 *
 * Usage:
 *   node scripts/capture-graph-fixtures.mjs
 *
 * Opens a dedicated Playwright Chromium browser window at http://127.0.0.1:8788/
 * allowing a developer to sign in with Consumer (OneDrive) and/or Enterprise
 * (Microsoft Entra ID / SharePoint / Teams) credentials.
 *
 * Intercepts the `obsidian://entra-sync-auth` PKCE redirect via Chrome DevTools Protocol,
 * exercises Microsoft Graph endpoints inside a self-cleaning `_ObsidianSyncFixtureProbe`
 * temporary folder, and runs a multi-layer deterministic anonymizer in memory before
 * writing sanitized fixtures to `test/fixtures/graph-<edition>.sanitized.json`.
 */

import http from "node:http";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const OUT_DIR = path.join(REPO_ROOT, "test", "fixtures");
const REDIRECT_URI = "obsidian://entra-sync-auth";
const DEFAULT_CONSUMER_CLIENT_ID = "ada87c12-a404-484b-a7dc-10c54beb2d7d";
const PORT = Number(process.env.CAPTURE_PORT || 8788);

const SAFE_PROBE_NAMES = new Set([
  "root",
  "Documents",
  "Documenten",
  "Shared Documents",
  "Gedeelde Documenten",
  "_ObsidianSyncFixtureProbe",
  "TestVault",
  "Notes",
  "Hello.md",
]);

const PLATFORM_HOSTS = new Set([
  "graph.microsoft.com",
  "login.microsoftonline.com",
  "login.live.com",
  "1drv.ms",
  "1drv.com",
  "onedrive.live.com",
  "my.microsoftpersonalcontent.com",
]);

function base64Url(buf) {
  return buf
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

/**
 * Deterministic in-memory structural anonymizer.
 * Preserves exact Microsoft Graph JSON structure, status codes, Drive/Item ID formats
 * (UUID vs SharePoint `b!...` vs `01...` Base32 item IDs vs Consumer `CID!s...` vs
 * SharePoint `"{GUID},ver"` ETags vs Base64 `permission.id` claims),
 * while scrubbing 100% of tokens, emails, names, tenant/client IDs, domains, quotas, and user folder names.
 */
export function createAnonymizer(mode, explicitSecrets = []) {
  const uuidMap = new Map();
  const bDriveMap = new Map();
  const cidMap = new Map();
  const itemSuffixMap = new Map();
  const spItemMap = new Map();
  const emailMap = new Map();
  const hostMap = new Map();
  const nameMap = new Map();
  const literalReplacements = new Map();

  let uuidSeq = 1;
  let bDriveSeq = 1;
  let cidSeq = 1;
  let itemSuffixSeq = 1;
  let spItemSeq = 1;
  let emailSeq = 1;
  let hostSeq = 1;
  let folderSeq = 1;

  const syntheticTenantDomain =
    mode === "consumer" ? "outlook.example.com" : "contoso.example.com";
  const syntheticPrimaryName = mode === "consumer" ? "Alex Rivera" : "Jordan Patel";

  function addLiteral(raw, replacement) {
    if (!raw || typeof raw !== "string") return;
    const trimmed = raw.trim();
    if (trimmed.length < 3) return;
    literalReplacements.set(trimmed, replacement);
    literalReplacements.set(trimmed.toLowerCase(), replacement);
    literalReplacements.set(encodeURIComponent(trimmed), encodeURIComponent(replacement));
  }

  for (const [raw, rep] of explicitSecrets) {
    addLiteral(raw, rep);
  }

  function anonUuid(u) {
    const key = u.toLowerCase();
    if (!uuidMap.has(key)) {
      const hex = String(uuidSeq++).padStart(12, "0");
      uuidMap.set(key, `00000000-0000-4000-8000-${hex}`);
    }
    const mapped = uuidMap.get(key);
    return u === u.toUpperCase() ? mapped.toUpperCase() : mapped;
  }

  function anonBDrive(bId) {
    const isEncoded = bId.toLowerCase().startsWith("b%21");
    const rawKey = bId.replace(/^b(?:!|%21)/i, "");
    if (!bDriveMap.has(rawKey)) {
      bDriveMap.set(
        rawKey,
        `SyntheticSharePointDriveId${String(bDriveSeq++).padStart(4, "0")}00000000000000000000`
      );
    }
    return `${isEncoded ? "b%21" : "b!"}${bDriveMap.get(rawKey)}`;
  }

  function anonSpItemId(spId) {
    if (!spItemMap.has(spId)) {
      spItemMap.set(
        spId,
        `01SYNTHETICITEM${String(spItemSeq++).padStart(19, "0")}`
      );
    }
    return spItemMap.get(spId);
  }

  function anonHexCid(cidPart) {
    const key = cidPart.toUpperCase();
    if (!cidMap.has(key)) {
      cidMap.set(key, `000000000000${String(cidSeq++).padStart(4, "0")}`);
    }
    return cidMap.get(key);
  }

  function anonCidItem(cidExpr) {
    const isEncoded = /%21/i.test(cidExpr);
    const sep = isEncoded ? /%21/i : "!";
    const [cidPart, suffixPart] = cidExpr.split(sep);
    const mappedCid = anonHexCid(cidPart);
    if (suffixPart === undefined) return mappedCid;

    let mappedSuffix = suffixPart;
    if (!/^\d+$/.test(suffixPart)) {
      if (!itemSuffixMap.has(suffixPart)) {
        itemSuffixMap.set(
          suffixPart,
          `s${String(itemSuffixSeq++).padStart(32, "0")}`
        );
      }
      mappedSuffix = itemSuffixMap.get(suffixPart);
    }
    return `${mappedCid}${isEncoded ? "%21" : "!"}${mappedSuffix}`;
  }

  function anonEmail(email) {
    const lower = email.toLowerCase();
    if (/^user\d+@(?:outlook|contoso)\.example\.com$/.test(lower)) {
      return lower;
    }
    if (!emailMap.has(lower)) {
      emailMap.set(lower, `user${emailSeq++}@${syntheticTenantDomain}`);
    }
    return emailMap.get(lower);
  }

  function anonHost(hostname) {
    const lower = hostname.toLowerCase();
    if (PLATFORM_HOSTS.has(lower) || lower.endsWith(".example.com")) {
      return lower;
    }
    if (!hostMap.has(lower)) {
      if (lower.endsWith("-my.sharepoint.com")) {
        hostMap.set(lower, "contoso-my.sharepoint.com");
      } else if (lower.endsWith(".sharepoint.com")) {
        hostMap.set(lower, "contoso.sharepoint.com");
      } else if (lower.endsWith(".onmicrosoft.com")) {
        hostMap.set(lower, "contoso.onmicrosoft.com");
      } else {
        hostMap.set(lower, `tenant${hostSeq++}.example.com`);
      }
    }
    return hostMap.get(lower);
  }

  function anonFolderOrSiteName(rawName, prefix = "SampleItem") {
    if (!rawName || typeof rawName !== "string") return rawName;
    if (
      SAFE_PROBE_NAMES.has(rawName) ||
      rawName.startsWith("_ObsidianSyncFixtureProbe") ||
      /^Sample(?:Site|Folder|Item)_\d+$/.test(rawName)
    ) {
      return rawName.replace(
        /_ObsidianSyncFixtureProbe_[a-zA-Z0-9_-]+/g,
        "_ObsidianSyncFixtureProbe"
      );
    }
    if (!nameMap.has(rawName)) {
      const mapped = `${prefix}_${folderSeq++}`;
      nameMap.set(rawName, mapped);
      addLiteral(rawName, mapped);
    }
    return nameMap.get(rawName);
  }

  function maybeScrubBase64Claim(str) {
    if (!/^[A-Za-z0-9+/=_-]{24,}$/.test(str) || /^01[A-Z0-9]{32}$/.test(str)) {
      return null;
    }
    try {
      const decoded = Buffer.from(str, "base64").toString("utf8");
      if (
        decoded &&
        /^[\x20-\x7E]+$/.test(decoded) &&
        (decoded.includes("membership|") ||
          decoded.includes("@") ||
          decoded.includes("sharepoint.com") ||
          decoded.includes("1drv.ms"))
      ) {
        const scrubbedDecoded = scrubString(decoded, "");
        return Buffer.from(scrubbedDecoded, "utf8").toString("base64");
      }
    } catch {
      // ignore non-base64
    }
    return null;
  }

  function scrubString(input, fieldKey = "") {
    if (typeof input !== "string" || input.length === 0) return input;
    let s = input;

    if (
      fieldKey === "access_token" ||
      fieldKey === "refresh_token" ||
      fieldKey === "id_token"
    ) {
      return `<REDACTED_${fieldKey.toUpperCase()}>`;
    }

    if (
      fieldKey === "@microsoft.graph.downloadUrl" ||
      fieldKey === "@microsoft.graph.downloadUrlNoAuth"
    ) {
      try {
        const u = new URL(s);
        return `https://${anonHost(u.hostname)}/redacted-download-url`;
      } catch {
        return "https://contoso.sharepoint.com/redacted-download-url";
      }
    }

    if (fieldKey === "hostname") {
      return anonHost(s);
    }

    // Normalize ISO timestamps so exact account/session creation times are not fingerprinted
    if (
      /DateTime$/i.test(fieldKey) ||
      (fieldKey === "date" && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(s))
    ) {
      return s.endsWith("Z") ? "2026-01-15T10:00:00Z" : "2026-01-15T10:00:00";
    }

    // Base64-encoded SharePoint permission IDs / shareIds (e.g. "i:0#.f|membership|user@domain")
    const scrubbedB64 = maybeScrubBase64Claim(s);
    if (scrubbedB64) {
      return scrubbedB64;
    }

    // Normalize dynamic probe folder suffix
    s = s.replace(/_ObsidianSyncFixtureProbe_[a-zA-Z0-9_-]+/g, "_ObsidianSyncFixtureProbe");

    // Apply known literal replacements (longest first)
    const sortedLiterals = Array.from(literalReplacements.entries()).sort(
      (a, b) => b[0].length - a[0].length
    );
    for (const [needle, rep] of sortedLiterals) {
      if (needle && s.includes(needle)) {
        s = s.split(needle).join(rep);
      }
    }

    // Emails / UPNs (both plain and %40 URL-encoded)
    s = s.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, (m) => anonEmail(m));
    s = s.replace(
      /([a-zA-Z0-9._%+-]+)%40([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/gi,
      (_, local, dom) => encodeURIComponent(anonEmail(`${local}@${dom}`))
    );

    // SharePoint b! / b%21 drive IDs
    s = s.replace(/\bb(?:!|%21)[A-Za-z0-9_-]{12,}/gi, (m) => anonBDrive(m));

    // Share tokens u! / u%21
    s = s.replace(
      /\bu(!|%21)[A-Za-z0-9_-]{12,}/gi,
      (_, bang) =>
        `u${bang}aHR0cHM6Ly9jb250b3NvLnNoYXJlcG9pbnQuY29tL3NpdGVzL1NhbXBsZVNpdGVfMQ`
    );

    // Consumer OneDrive CID!itemId or CID%21itemId patterns
    s = s.replace(
      /\b[0-9A-Fa-f]{12,20}(?:!|%21)[A-Za-z0-9._-]+\b/g,
      (m) => anonCidItem(m)
    );

    // Standalone 16-hex Consumer CIDs
    s = s.replace(/\b[0-9A-Fa-f]{16}\b/g, (m) => anonHexCid(m));

    // SharePoint 34-char Base32 DriveItem IDs (01[A-Z2-7]{32})
    s = s.replace(/\b01[A-Z2-7]{32}\b/g, (m) => anonSpItemId(m));

    // UUIDs / GUIDs (including inside "{GUID},1" SharePoint ETags!)
    s = s.replace(
      /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g,
      (m) => anonUuid(m)
    );

    // Normalize 1drv.ms short links
    s = s.replace(
      /https:\/\/1drv\.ms\/[a-z]\/c\/[^/]+\/[^"'\s)]+/gi,
      "https://1drv.ms/f/c/0000000000000001/SyntheticOneDriveShareToken"
    );

    // Scrub URLs, hostnames, /personal/<slug>, /sites/<slug>, /teams/<slug>, delta tokens, and query params
    if (/^https?:\/\//i.test(s)) {
      try {
        const u = new URL(s);
        u.hostname = anonHost(u.hostname);
        u.pathname = u.pathname
          .replace(/\/personal\/[^/]+/gi, "/personal/user1_contoso_example_com")
          .replace(/\/(sites|teams)\/([^/]+)/gi, (_, kind, slug) => {
            const safeSlug = anonFolderOrSiteName(decodeURIComponent(slug), "SampleSite");
            return `/${kind}/${encodeURIComponent(safeSlug)}`;
          })
          .replace(/\/Gedeelde%20SampleFolder_\d+/gi, "/Gedeelde%20Documenten");
        for (const [k] of Array.from(u.searchParams.entries())) {
          if (/token|skiptoken|authkey|tempauth|sig|resid|cid|id|sourcedoc|uniqueid/i.test(k)) {
            u.searchParams.set(k, `REDACTED_${k.toUpperCase()}`);
          }
        }
        s = u.toString();
      } catch {
        // leave scrubbed string
      }
    }

    // Bare *.sharepoint.com / *.onmicrosoft.com domains anywhere in non-URL strings
    s = s.replace(
      /\b([a-zA-Z0-9-]+(?:-my)?\.sharepoint\.com)\b/gi,
      (m) => anonHost(m)
    );
    s = s.replace(
      /\b([a-zA-Z0-9-]+\.onmicrosoft\.com)\b/gi,
      (m) => anonHost(m)
    );

    // Normalize @microsoft.graph.lastProcessedChange sequence numbers
    if (fieldKey === "@microsoft.graph.lastProcessedChange") {
      return "1;1;00000000-0000-4000-8000-000000000006;639000000000000000;1000000001";
    }

    return s;
  }

  function registerProfileIdentity(meJson, driveJson) {
    if (meJson && typeof meJson === "object") {
      if (meJson.displayName) addLiteral(meJson.displayName, syntheticPrimaryName);
      if (meJson.givenName) addLiteral(meJson.givenName, syntheticPrimaryName.split(" ")[0]);
      if (meJson.surname) addLiteral(meJson.surname, syntheticPrimaryName.split(" ")[1]);
      if (meJson.mail) addLiteral(meJson.mail, anonEmail(meJson.mail));
      if (meJson.userPrincipalName) {
        addLiteral(meJson.userPrincipalName, anonEmail(meJson.userPrincipalName));
      }
      if (meJson.id && /^[0-9a-fA-F]{12,20}$/.test(meJson.id)) {
        addLiteral(meJson.id, anonHexCid(meJson.id));
      }
    }
    if (driveJson && typeof driveJson === "object") {
      if (typeof driveJson.id === "string" && /^[0-9a-fA-F]{12,20}$/.test(driveJson.id)) {
        addLiteral(driveJson.id, anonHexCid(driveJson.id));
      }
      const ownerUser = driveJson.owner?.user;
      if (ownerUser?.displayName) addLiteral(ownerUser.displayName, syntheticPrimaryName);
      if (ownerUser?.email) addLiteral(ownerUser.email, anonEmail(ownerUser.email));
    }
  }

  function scrubValue(val, key = "", context = { anonymizeItemNames: false }) {
    if (val === null || val === undefined) return val;
    if (typeof val === "string") {
      if (
        context.anonymizeItemNames &&
        (key === "name" || key === "displayName") &&
        !SAFE_PROBE_NAMES.has(val) &&
        !val.startsWith("_ObsidianSyncFixtureProbe")
      ) {
        return anonFolderOrSiteName(val, key === "displayName" ? "SampleSite" : "SampleFolder");
      }
      if (key === "surname") return syntheticPrimaryName.split(" ")[1];
      if (key === "givenName") return syntheticPrimaryName.split(" ")[0];
      if (key === "jobTitle" || key === "officeLocation" || key === "mobilePhone") return null;
      return scrubString(val, key);
    }
    if (Array.isArray(val)) {
      if (key === "businessPhones") return [];
      return val.map((item) => scrubValue(item, key, context));
    }
    if (typeof val === "object") {
      // Normalize drive storage quota numbers so exact disk usage cannot fingerprint a tenant/user
      if (key === "quota" && "total" in val && "remaining" in val) {
        const normalizedQuota = {
          deleted: 0,
          remaining: 1099406770176,
          state: val.state || "normal",
          total: 1099511627776,
          used: 104857600,
        };
        if ("storagePlanInformation" in val) {
          normalizedQuota.storagePlanInformation = { upgradeAvailable: false };
        }
        return normalizedQuota;
      }

      const out = {};
      const isUserNode =
        Boolean(val.displayName) &&
        ("email" in val || "id" in val) &&
        !("folder" in val) &&
        !("file" in val) &&
        !("driveType" in val) &&
        !("webUrl" in val) &&
        !("siteCollection" in val) &&
        !("root" in val);
      for (const [k, v] of Object.entries(val)) {
        if (isUserNode && k === "displayName" && typeof v === "string") {
          if (
            v !== syntheticPrimaryName &&
            !v.startsWith("Collaborator ") &&
            !literalReplacements.has(v)
          ) {
            addLiteral(v, `Collaborator ${folderSeq++}`);
          }
          out[k] = scrubString(v, k);
        } else {
          out[k] = scrubValue(v, k, context);
        }
      }
      return out;
    }
    return val;
  }

  return {
    registerProfileIdentity,
    scrubString,
    scrubPayload(payload, { anonymizeItemNames = false } = {}) {
      return scrubValue(payload, "", { anonymizeItemNames });
    },
  };
}

async function graphCall(accessToken, method, url, { headers = {}, body } = {}) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...headers,
    },
    body,
  });
  const text = await res.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = { _nonJsonText: text.slice(0, 200) };
  }
  const relevantHeaders = {};
  for (const h of ["content-type", "etag", "location", "request-id"]) {
    if (res.headers.has(h)) {
      relevantHeaders[h] =
        h === "request-id"
          ? "00000000-0000-4000-8000-000000000000"
          : res.headers.get(h);
    }
  }
  return { status: res.status, headers: relevantHeaders, body: json };
}

async function runGraphProbeSession({
  mode,
  tenantId,
  clientId,
  code,
  codeVerifier,
  sharePointProbeUrl = "",
  onProgress,
}) {
  const effectiveTenant = mode === "consumer" ? "common" : tenantId.trim();
  const effectiveClient = mode === "consumer" ? DEFAULT_CONSUMER_CLIENT_ID : clientId.trim();

  const anonymizer = createAnonymizer(mode, [
    [effectiveTenant !== "common" ? effectiveTenant : "", "00000000-0000-4000-8000-000000000999"],
    [
      effectiveClient !== DEFAULT_CONSUMER_CLIENT_ID ? effectiveClient : "",
      "00000000-0000-4000-8000-000000000888",
    ],
  ]);

  onProgress("Exchanging PKCE authorization code with Microsoft identity platform...");
  const tokenParams = new URLSearchParams({
    client_id: effectiveClient,
    scope:
      mode === "enterprise"
        ? "openid profile offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All"
        : "openid profile offline_access User.Read Files.ReadWrite.All",
    code,
    redirect_uri: REDIRECT_URI,
    grant_type: "authorization_code",
    code_verifier: codeVerifier,
  });

  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(effectiveTenant)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: tokenParams.toString(),
    }
  );
  const tokenJson = await tokenRes.json();
  if (!tokenRes.ok || !tokenJson.access_token) {
    throw new Error(
      `Token exchange failed (${tokenRes.status}): ${tokenJson.error || "unknown"} - ${
        tokenJson.error_description || ""
      }`
    );
  }

  let accessToken = tokenJson.access_token;
  const refreshToken = tokenJson.refresh_token;

  let refreshResponseSample = null;
  if (refreshToken) {
    onProgress("Testing OAuth2 PKCE refresh_token rotation...");
    const refParams = new URLSearchParams({
      client_id: effectiveClient,
      scope:
        mode === "enterprise"
          ? "openid profile offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All"
          : "openid profile offline_access User.Read Files.ReadWrite.All",
      refresh_token: refreshToken,
      grant_type: "refresh_token",
    });
    const refRes = await fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(effectiveTenant)}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: refParams.toString(),
      }
    );
    const refJson = await refRes.json();
    if (refRes.ok && refJson.access_token) {
      accessToken = refJson.access_token;
    }
    refreshResponseSample = {
      status: refRes.status,
      body: anonymizer.scrubPayload(refJson),
    };
  }

  const base = "https://graph.microsoft.com/v1.0";
  const probeRootName = `_ObsidianSyncFixtureProbe_${crypto.randomBytes(4).toString("hex")}`;
  let probeRootItemId = null;
  let myDriveId = null;

  const captured = {
    capturedAt: "2026-01-15",
    edition: mode,
    oauthTokenExchange: {
      status: tokenRes.status,
      body: anonymizer.scrubPayload(tokenJson),
    },
    oauthTokenRefresh: refreshResponseSample,
    endpoints: {},
  };

  try {
    onProgress("Probing /me and /me/drive...");
    const meCall = await graphCall(accessToken, "GET", `${base}/me`);
    const myDriveCall = await graphCall(accessToken, "GET", `${base}/me/drive`);
    anonymizer.registerProfileIdentity(meCall.body, myDriveCall.body);

    captured.endpoints["GET /v1.0/me"] = {
      status: meCall.status,
      body: anonymizer.scrubPayload(meCall.body),
    };
    captured.endpoints["GET /v1.0/me/drive"] = {
      status: myDriveCall.status,
      body: anonymizer.scrubPayload(myDriveCall.body),
    };
    myDriveId = myDriveCall.body?.id;

    onProgress("Creating temporary isolated probe folder in OneDrive (deleted automatically)...");
    const createRootProbe = await graphCall(accessToken, "POST", `${base}/me/drive/root/children`, {
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: probeRootName, folder: {} }),
    });
    probeRootItemId = createRootProbe.body?.id;
    myDriveId = myDriveId || createRootProbe.body?.parentReference?.driveId;

    captured.endpoints["POST /v1.0/me/drive/root/children (create parent folder)"] = {
      status: createRootProbe.status,
      body: anonymizer.scrubPayload(createRootProbe.body),
    };

    if (probeRootItemId && myDriveId) {
      const createVaultCall = await graphCall(
        accessToken,
        "POST",
        `${base}/me/drive/items/${probeRootItemId}/children`,
        {
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ name: "TestVault", folder: {} }),
        }
      );
      const vaultItemId = createVaultCall.body?.id;
      captured.endpoints["POST /v1.0/me/drive/items/{parentId}/children (create vault)"] = {
        status: createVaultCall.status,
        body: anonymizer.scrubPayload(createVaultCall.body),
      };

      const duplicateVaultCall = await graphCall(
        accessToken,
        "POST",
        `${base}/me/drive/items/${probeRootItemId}/children`,
        {
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            name: "TestVault",
            folder: {},
            "@microsoft.graph.conflictBehavior": "fail",
          }),
        }
      );
      captured.endpoints["POST /v1.0/me/drive/items/{parentId}/children (duplicate folder 409)"] = {
        status: duplicateVaultCall.status,
        body: anonymizer.scrubPayload(duplicateVaultCall.body),
      };

      const listChildrenCall = await graphCall(
        accessToken,
        "GET",
        `${base}/me/drive/root:/${encodeURIComponent(probeRootName)}:/children`
      );
      captured.endpoints["GET /v1.0/me/drive/root:/{folder}:/children"] = {
        status: listChildrenCall.status,
        body: anonymizer.scrubPayload(listChildrenCall.body),
      };

      if (vaultItemId) {
        onProgress("Probing /delta, PUT content, ETag 412 conflict, /versions, /permissions, and /shares...");
        const initialDeltaCall = await graphCall(
          accessToken,
          "GET",
          `${base}/drives/${myDriveId}/items/${vaultItemId}/delta`
        );
        const deltaLink1 = initialDeltaCall.body?.["@odata.deltaLink"];
        captured.endpoints["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (initial empty)"] = {
          status: initialDeltaCall.status,
          body: anonymizer.scrubPayload(initialDeltaCall.body),
        };

        const putV1Call = await graphCall(
          accessToken,
          "PUT",
          `${base}/drives/${myDriveId}/items/${vaultItemId}:/Notes/Hello.md:/content`,
          {
            headers: { "Content-Type": "text/markdown; charset=utf-8" },
            body: "# Hello World\nVersion 1",
          }
        );
        const fileId = putV1Call.body?.id;
        const v1ETag = putV1Call.body?.eTag;
        captured.endpoints["PUT /v1.0/drives/{driveId}/items/{vaultId}:/{path}:/content (create v1)"] = {
          status: putV1Call.status,
          headers: anonymizer.scrubPayload(putV1Call.headers),
          body: anonymizer.scrubPayload(putV1Call.body),
        };

        if (v1ETag) {
          const putV2Call = await graphCall(
            accessToken,
            "PUT",
            `${base}/drives/${myDriveId}/items/${vaultItemId}:/Notes/Hello.md:/content`,
            {
              headers: {
                "Content-Type": "text/markdown; charset=utf-8",
                "If-Match": v1ETag,
              },
              body: "# Hello World\nVersion 2",
            }
          );
          captured.endpoints["PUT /v1.0/drives/{driveId}/items/{vaultId}:/{path}:/content (update v2 with If-Match)"] = {
            status: putV2Call.status,
            headers: anonymizer.scrubPayload(putV2Call.headers),
            body: anonymizer.scrubPayload(putV2Call.body),
          };

          const putStaleCall = await graphCall(
            accessToken,
            "PUT",
            `${base}/drives/${myDriveId}/items/${vaultItemId}:/Notes/Hello.md:/content`,
            {
              headers: {
                "Content-Type": "text/markdown; charset=utf-8",
                "If-Match": v1ETag,
              },
              body: "# Concurrent Stale Edit",
            }
          );
          captured.endpoints["PUT /v1.0/drives/{driveId}/items/{vaultId}:/{path}:/content (stale If-Match -> 412)"] = {
            status: putStaleCall.status,
            headers: anonymizer.scrubPayload(putStaleCall.headers),
            body: anonymizer.scrubPayload(putStaleCall.body),
          };
        }

        if (fileId) {
          const itemMetaCall = await graphCall(
            accessToken,
            "GET",
            `${base}/drives/${myDriveId}/items/${fileId}`
          );
          captured.endpoints["GET /v1.0/drives/{driveId}/items/{itemId}"] = {
            status: itemMetaCall.status,
            body: anonymizer.scrubPayload(itemMetaCall.body),
          };

          const versionsCall = await graphCall(
            accessToken,
            "GET",
            `${base}/drives/${myDriveId}/items/${fileId}/versions`
          );
          captured.endpoints["GET /v1.0/drives/{driveId}/items/{itemId}/versions"] = {
            status: versionsCall.status,
            body: anonymizer.scrubPayload(versionsCall.body),
          };
        }

        const permsCall = await graphCall(
          accessToken,
          "GET",
          `${base}/drives/${myDriveId}/items/${vaultItemId}/permissions`
        );
        captured.endpoints["GET /v1.0/drives/{driveId}/items/{vaultId}/permissions"] = {
          status: permsCall.status,
          body: anonymizer.scrubPayload(permsCall.body),
        };

        const linkScope = mode === "enterprise" ? "organization" : "anonymous";
        const createLinkCall = await graphCall(
          accessToken,
          "POST",
          `${base}/drives/${myDriveId}/items/${vaultItemId}/createLink`,
          {
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ type: "view", scope: linkScope }),
          }
        );
        const webShareUrl = createLinkCall.body?.link?.webUrl || createVaultCall.body?.webUrl;
        if (webShareUrl) {
          const shareToken = `u!${base64Url(Buffer.from(webShareUrl, "utf8"))}`;
          const shareResolveCall = await graphCall(
            accessToken,
            "GET",
            `${base}/shares/${shareToken}/driveItem`
          );
          captured.endpoints["GET /v1.0/shares/{u!token}/driveItem (TestVault URL)"] = {
            status: shareResolveCall.status,
            body: anonymizer.scrubPayload(shareResolveCall.body),
          };
        }

        if (deltaLink1) {
          const deltaChangedCall = await graphCall(accessToken, "GET", deltaLink1);
          const deltaLink2 = deltaChangedCall.body?.["@odata.deltaLink"];
          captured.endpoints["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (incremental create)"] = {
            status: deltaChangedCall.status,
            body: anonymizer.scrubPayload(deltaChangedCall.body),
          };

          if (fileId && deltaLink2) {
            await graphCall(accessToken, "DELETE", `${base}/drives/${myDriveId}/items/${fileId}`);
            const deltaDeletedCall = await graphCall(accessToken, "GET", deltaLink2);
            captured.endpoints["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (incremental delete tombstone)"] = {
              status: deltaDeletedCall.status,
              body: anonymizer.scrubPayload(deltaDeletedCall.body),
            };
          }
        }
      }
    }

    onProgress("Probing discovery endpoints (sharedWithMe & followedSites, strictly anonymized)...");
    const sharedCall = await graphCall(
      accessToken,
      "GET",
      `${base}/me/drive/sharedWithMe?$top=2`
    );
    captured.endpoints["GET /v1.0/me/drive/sharedWithMe"] = {
      status: sharedCall.status,
      body: anonymizer.scrubPayload(sharedCall.body, { anonymizeItemNames: true }),
    };

    if (mode === "enterprise") {
      const sitesCall = await graphCall(
        accessToken,
        "GET",
        `${base}/me/followedSites?$top=2`
      );
      captured.endpoints["GET /v1.0/me/followedSites"] = {
        status: sitesCall.status,
        body: anonymizer.scrubPayload(sitesCall.body, { anonymizeItemNames: true }),
      };

      const firstSite = sitesCall.body?.value?.[0];
      if (firstSite?.id) {
        const siteDriveCall = await graphCall(
          accessToken,
          "GET",
          `${base}/sites/${firstSite.id}/drive`
        );
        captured.endpoints["GET /v1.0/sites/{siteId}/drive"] = {
          status: siteDriveCall.status,
          body: anonymizer.scrubPayload(siteDriveCall.body, { anonymizeItemNames: true }),
        };

        const spDriveId = siteDriveCall.body?.id;
        if (spDriveId) {
          const spRootChildren = await graphCall(
            accessToken,
            "GET",
            `${base}/drives/${spDriveId}/root/children?$top=2`
          );
          captured.endpoints["GET /v1.0/drives/{spDriveId}/root/children"] = {
            status: spRootChildren.status,
            body: anonymizer.scrubPayload(spRootChildren.body, { anonymizeItemNames: true }),
          };
        }
      }

      if (sharePointProbeUrl.trim()) {
        onProgress("Probing provided SharePoint/Teams URL (read-only, strictly anonymized)...");
        let normalizedUrl = sharePointProbeUrl.trim();
        try {
          const u = new URL(normalizedUrl);
          const idParam = u.searchParams.get("id");
          if (idParam && idParam.startsWith("/")) {
            normalizedUrl = `${u.origin}${idParam
              .split("/")
              .map((seg) => encodeURIComponent(decodeURIComponent(seg)))
              .join("/")}`;
          }
        } catch {
          // keep raw
        }
        const spShareToken = `u!${base64Url(Buffer.from(normalizedUrl, "utf8"))}`;
        const spShareCall = await graphCall(
          accessToken,
          "GET",
          `${base}/shares/${spShareToken}/driveItem`
        );
        captured.endpoints["GET /v1.0/shares/{u!token}/driveItem (SharePoint/Teams URL)"] = {
          status: spShareCall.status,
          body: anonymizer.scrubPayload(spShareCall.body, { anonymizeItemNames: true }),
        };
      }
    }
  } finally {
    if (probeRootItemId && myDriveId) {
      onProgress("Cleaning up temporary probe folder from OneDrive...");
      try {
        await fetch(`${base}/drives/${myDriveId}/items/${probeRootItemId}`, {
          method: "DELETE",
          headers: { Authorization: `Bearer ${accessToken}` },
        });
      } catch {
        // ignore cleanup error
      }
    }
    accessToken = "";
  }

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const outPath = path.join(OUT_DIR, `graph-${mode}.sanitized.json`);
  fs.writeFileSync(outPath, JSON.stringify(captured, null, 2) + "\n", "utf8");
  onProgress(`Done! Saved anonymized ${mode} fixture (${Object.keys(captured.endpoints).length} endpoints).`);
  return captured;
}

const HTML_UI = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Obsidian Sync — Anonymized Microsoft Graph Fixture Capture</title>
  <style>
    :root { color-scheme: dark; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    body { max-width: 880px; margin: 32px auto; padding: 0 20px; background: #16171a; color: #e5e7eb; }
    h1 { font-size: 22px; margin-bottom: 6px; }
    p.lead { color: #9ca3af; font-size: 14px; margin-top: 0; line-height: 1.5; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-top: 20px; }
    .card { background: #1f2128; border: 1px solid #323642; border-radius: 10px; padding: 18px; display: flex; flex-direction: column; gap: 12px; }
    .card h2 { font-size: 16px; margin: 0; }
    .muted { color: #9ca3af; font-size: 13px; margin: 0; line-height: 1.45; }
    label { display: block; font-size: 12px; font-weight: 600; color: #d1d5db; margin-bottom: 4px; }
    input { width: 100%; box-sizing: border-box; padding: 8px 10px; border-radius: 6px; border: 1px solid #3f4454; background: #14161b; color: #f3f4f6; font-size: 13px; }
    button { padding: 9px 14px; border-radius: 6px; border: none; background: #7c3aed; color: white; font-weight: 600; font-size: 13px; cursor: pointer; }
    button:hover { background: #6d28d9; }
    .status { padding: 8px 10px; border-radius: 6px; font-size: 12.5px; background: #14161b; border: 1px solid #2d313d; }
    .status.done { border-color: #10b981; color: #6ee7b7; }
    .status.error { border-color: #ef4444; color: #fca5a5; }
    .status.busy { border-color: #f59e0b; color: #fcd34d; }
    pre { background: #0f1115; border: 1px solid #2b2f3a; border-radius: 8px; padding: 12px; font-size: 11.5px; max-height: 240px; overflow: auto; white-space: pre-wrap; word-break: break-all; margin: 0; }
    .privacy-box { margin-top: 18px; padding: 12px 14px; border-radius: 8px; background: rgba(16,185,129,0.08); border: 1px solid rgba(16,185,129,0.28); font-size: 12.5px; color: #a7f3d0; line-height: 1.5; }
  </style>
</head>
<body>
  <h1>Microsoft Graph Anonymized Fixture Capture</h1>
  <p class="lead">
    Sign in below with either or both accounts. A login tab will open right inside this browser, intercept the
    <code>obsidian://entra-sync-auth</code> PKCE callback automatically, run a safe temporary folder probe (cleaned up automatically),
    and scrub <strong>all tokens, emails, names, tenant/client IDs, domains, drive IDs, and folder names</strong> in memory before saving.
  </p>

  <div class="grid">
    <div class="card">
      <h2>1. Consumer (OneDrive)</h2>
      <p class="muted">Personal Microsoft account (<code>outlook.com</code>, <code>live.com</code>, <code>hotmail.com</code>). Uses <code>common</code> + public client PKCE.</p>
      <div style="flex:1"></div>
      <button onclick="startCapture('consumer')">Sign in with Consumer Account</button>
      <div id="status-consumer" class="status">Ready to sign in</div>
      <pre id="preview-consumer" style="display:none"></pre>
    </div>

    <div class="card">
      <h2>2. Enterprise (Teams &amp; SharePoint)</h2>
      <p class="muted">Corporate Microsoft Entra ID account. Requests <code>User.Read</code>, <code>Files.ReadWrite.All</code>, and <code>Sites.ReadWrite.All</code>.</p>
      <div>
        <label>Directory (tenant) ID</label>
        <input id="ent-tenant" placeholder="e.g. contoso.onmicrosoft.com or tenant UUID (or common)" />
      </div>
      <div>
        <label>Application (client) ID</label>
        <input id="ent-client" value="${DEFAULT_CONSUMER_CLIENT_ID}" placeholder="Application (client) ID" />
      </div>
      <div>
        <label>Optional SharePoint / Teams Vault URL to probe</label>
        <input id="ent-sp-url" placeholder="https://contoso.sharepoint.com/sites/Team/Shared%20Documents/..." />
      </div>
      <button onclick="startCapture('enterprise')">Sign in with Enterprise Account</button>
      <div id="status-enterprise" class="status">Ready to sign in</div>
      <pre id="preview-enterprise" style="display:none"></pre>
    </div>
  </div>

  <div class="privacy-box">
    🔒 <strong>Privacy guarantee:</strong> Access &amp; refresh tokens exist only in RAM during the 5-second probe and are immediately wiped.
    All emails, names, tenant/client UUIDs, SharePoint hostnames, OneDrive CIDs, Base64 permission claims, quotas, and folder/site names are replaced with synthetic placeholders before writing JSON to <code>test/fixtures/</code>.
  </div>

  <script>
    async function startCapture(mode) {
      const payload = { mode };
      if (mode === 'enterprise') {
        payload.tenantId = document.getElementById('ent-tenant').value.trim() || 'common';
        payload.clientId = document.getElementById('ent-client').value.trim();
        payload.sharePointProbeUrl = document.getElementById('ent-sp-url').value.trim();
      }
      await fetch('/api/start', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    }

    setInterval(async () => {
      try {
        const res = await fetch('/api/status');
        const data = await res.json();
        for (const mode of ['consumer', 'enterprise']) {
          const el = document.getElementById('status-' + mode);
          const pre = document.getElementById('preview-' + mode);
          const info = data[mode];
          el.className = 'status ' + info.state;
          el.textContent = info.message;
          if (info.preview) {
            pre.style.display = 'block';
            pre.textContent = info.preview;
          }
        }
      } catch {}
    }, 1000);
  </script>
</body>
</html>`;

async function startInteractiveCaptureServer() {
  const require = createRequire(path.join(REPO_ROOT, "package.json"));
  const { chromium } = require("playwright");

  const sessionStatus = {
    consumer: { state: "idle", message: "Ready to sign in", preview: null },
    enterprise: { state: "idle", message: "Ready to sign in", preview: null },
  };

  let browserContext = null;

  async function launchLoginFlow({
    mode,
    tenantId = "common",
    clientId = DEFAULT_CONSUMER_CLIENT_ID,
    sharePointProbeUrl = "",
  }) {
    sessionStatus[mode] = { state: "busy", message: "Waiting for browser sign-in...", preview: null };

    const verifier = base64Url(crypto.randomBytes(32));
    const challenge = base64Url(crypto.createHash("sha256").update(verifier).digest());
    const state = base64Url(crypto.randomBytes(16));
    const effectiveTenant = mode === "consumer" ? "common" : tenantId.trim();
    const effectiveClient = mode === "consumer" ? DEFAULT_CONSUMER_CLIENT_ID : clientId.trim();
    const scopes =
      mode === "enterprise"
        ? ["openid", "profile", "offline_access", "User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
        : ["openid", "profile", "offline_access", "User.Read", "Files.ReadWrite.All"];

    const params = new URLSearchParams({
      client_id: effectiveClient,
      response_type: "code",
      redirect_uri: REDIRECT_URI,
      response_mode: "query",
      scope: scopes.join(" "),
      code_challenge: challenge,
      code_challenge_method: "S256",
      state,
      prompt: "select_account",
    });
    const authUrl = `https://login.microsoftonline.com/${encodeURIComponent(effectiveTenant)}/oauth2/v2.0/authorize?${params.toString()}`;

    const loginPage = await browserContext.newPage();
    const cdp = await browserContext.newCDPSession(loginPage);
    await cdp.send("Network.enable");
    await cdp.send("Page.enable");

    let handled = false;
    const handleRedirectUrl = async (rawRedirectUrl) => {
      if (handled || !rawRedirectUrl.startsWith("obsidian://entra-sync-auth")) return;
      handled = true;
      try {
        const parsed = new URL(rawRedirectUrl);
        const err = parsed.searchParams.get("error");
        const errDesc = parsed.searchParams.get("error_description");
        const code = parsed.searchParams.get("code");
        await loginPage.close().catch(() => {});

        if (err || !code) {
          sessionStatus[mode] = {
            state: "error",
            message: `OAuth error: ${err || "missing_code"} ${errDesc || ""}`,
            preview: null,
          };
          return;
        }

        const result = await runGraphProbeSession({
          mode,
          tenantId: effectiveTenant,
          clientId: effectiveClient,
          code,
          codeVerifier: verifier,
          sharePointProbeUrl,
          onProgress: (msg) => {
            sessionStatus[mode] = { state: "busy", message: msg, preview: null };
          },
        });
        sessionStatus[mode] = {
          state: "done",
          message: `Captured & anonymized ${Object.keys(result.endpoints).length} Microsoft Graph responses!`,
          preview: JSON.stringify(result, null, 2).slice(0, 4000) + "\n... (truncated preview)",
        };
      } catch (e) {
        sessionStatus[mode] = {
          state: "error",
          message: String(e.message || e),
          preview: null,
        };
      }
    };

    cdp.on("Network.requestWillBeSent", (e) => {
      if (e.request?.url?.startsWith("obsidian://entra-sync-auth")) {
        void handleRedirectUrl(e.request.url);
      }
    });
    cdp.on("Network.responseReceivedExtraInfo", (e) => {
      const loc = e.headers?.location || e.headers?.Location;
      if (loc && loc.startsWith("obsidian://entra-sync-auth")) {
        void handleRedirectUrl(loc);
      }
    });
    cdp.on("Page.frameRequestedNavigation", (e) => {
      if (e.url && e.url.startsWith("obsidian://entra-sync-auth")) {
        void handleRedirectUrl(e.url);
      }
    });
    loginPage.on("close", () => {
      if (!handled && sessionStatus[mode].state === "busy") {
        sessionStatus[mode] = {
          state: "idle",
          message: "Sign-in window closed before completing.",
          preview: null,
        };
      }
    });

    await loginPage.goto(authUrl, { waitUntil: "domcontentloaded", timeout: 30000 }).catch(() => {});
  }

  const server = http.createServer(async (req, res) => {
    const u = new URL(req.url, `http://127.0.0.1:${PORT}`);
    if (req.method === "GET" && u.pathname === "/") {
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(HTML_UI);
      return;
    }
    if (req.method === "GET" && u.pathname === "/api/status") {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify(sessionStatus));
      return;
    }
    if (req.method === "POST" && u.pathname === "/api/start") {
      const chunks = [];
      for await (const c of req) chunks.push(c);
      const body = JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
      void launchLoginFlow(body);
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: true }));
      return;
    }
    res.writeHead(404);
    res.end("Not found");
  });

  await new Promise((resolve, reject) => {
    server.listen(PORT, "127.0.0.1", resolve);
    server.on("error", reject);
  });

  const browser = await chromium.launch({ headless: false });
  browserContext = await browser.newContext({ viewport: { width: 1020, height: 780 } });
  const dashboardPage = await browserContext.newPage();
  await dashboardPage.goto(`http://127.0.0.1:${PORT}/`);
  console.log(`Capture UI ready at http://127.0.0.1:${PORT}/`);

  browser.on("disconnected", () => {
    server.close();
    process.exit(0);
  });
}

if (import.meta.url === `file://${process.argv[1]}`) {
  await startInteractiveCaptureServer();
}
