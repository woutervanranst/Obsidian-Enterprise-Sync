import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import Module from "node:module";
import { fileURLToPath } from "node:url";
import esbuild from "esbuild";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const consumerFixturePath = path.join(__dirname, "fixtures", "graph-consumer.sanitized.json");
const enterpriseFixturePath = path.join(__dirname, "fixtures", "graph-enterprise.sanitized.json");

const consumerFixture = JSON.parse(fs.readFileSync(consumerFixturePath, "utf8"));
const enterpriseFixture = JSON.parse(fs.readFileSync(enterpriseFixturePath, "utf8"));

class MockTFile {
  constructor(filePath, content = "") {
    this.path = filePath;
    this.name = filePath.split("/").pop();
    this.extension = this.name.includes(".") ? this.name.split(".").pop() : "";
    this.stat = { mtime: Date.now(), ctime: Date.now(), size: content.length };
  }
}

class MockTFolder {
  constructor(folderPath) {
    this.path = folderPath;
    this.name = folderPath.split("/").pop();
  }
}

class MockVault {
  constructor() {
    this.configDir = ".obsidian";
    this.files = new Map();
    this.folders = new Map();
    this.adapter = {
      exists: async (p) => this.files.has(p) || this.folders.has(p),
      list: async () => ({ files: [], folders: [] }),
      read: async (p) => this.files.get(p)?.content ?? "",
      readBinary: async (p) => {
        const entry = this.files.get(p);
        const u8 = Buffer.from(entry?.content ?? "", "utf8");
        return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
      },
      write: async (p, content) => {
        await this.create(p, content);
      },
      mkdir: async (p) => {
        await this.createFolder(p);
      },
      remove: async (p) => {
        this.files.delete(p);
      },
      rmdir: async (p) => {
        this.folders.delete(p);
        const prefix = `${p}/`;
        for (const k of Array.from(this.files.keys())) {
          if (k.startsWith(prefix)) this.files.delete(k);
        }
        for (const k of Array.from(this.folders.keys())) {
          if (k.startsWith(prefix)) this.folders.delete(k);
        }
      },
    };
  }
  getFiles() {
    return Array.from(this.files.values()).map((f) => f.file);
  }
  getMarkdownFiles() {
    return Array.from(this.files.values())
      .filter((f) => f.file.extension === "md")
      .map((f) => f.file);
  }
  getAbstractFileByPath(p) {
    if (this.files.has(p)) return this.files.get(p).file;
    if (this.folders.has(p)) return this.folders.get(p);
    return null;
  }
  async read(tfile) {
    return this.files.get(tfile.path)?.content ?? "";
  }
  async create(p, content) {
    const tfile = new MockTFile(p, content);
    this.files.set(p, { file: tfile, content });
    return tfile;
  }
  async modify(tfile, content) {
    const entry = this.files.get(tfile.path);
    if (entry) entry.content = content;
    else this.files.set(tfile.path, { file: tfile, content });
  }
  async delete(abstractFile) {
    if (abstractFile instanceof MockTFolder) {
      await this.adapter.rmdir(abstractFile.path);
    } else {
      this.files.delete(abstractFile.path);
    }
  }
  async createFolder(p) {
    const folder = new MockTFolder(p);
    this.folders.set(p, folder);
    return folder;
  }
}

class MockElement {
  constructor(tag = "div", opts = {}) {
    this.tag = tag;
    this.cls = opts.cls || "";
    this.text = opts.text || "";
    this.children = [];
  }
  empty() {
    this.children = [];
  }
  addClass(cls) {
    this.cls = `${this.cls} ${cls}`.trim();
  }
  createEl(tag, opts = {}) {
    const child = new MockElement(tag, opts);
    this.children.push(child);
    return child;
  }
  createDiv(opts = {}) {
    return this.createEl("div", opts);
  }
  createSpan(opts = {}) {
    return this.createEl("span", opts);
  }
  findAllText() {
    const out = this.text ? [this.text] : [];
    for (const c of this.children) out.push(...c.findAllText());
    return out;
  }
}

// Active responder swapped per test
let activeHandler = async (opts) => ({ status: 404, headers: {}, text: "{}", json: {} });

function createObsidianShim() {
  class Modal {
    constructor(app) {
      this.app = app;
      this.contentEl = new MockElement("div");
    }
    open() {}
    close() {}
  }
  class Setting {
    constructor(containerEl) {
      this.containerEl = containerEl;
    }
    setName() { return this; }
    setDesc() { return this; }
    setHeading() { return this; }
    addText(cb) {
      cb({
        inputEl: {},
        setPlaceholder() { return this; },
        setValue() { return this; },
        onChange() { return this; },
      });
      return this;
    }
    addDropdown(cb) {
      cb({
        addOption() { return this; },
        setValue() { return this; },
        onChange() { return this; },
      });
      return this;
    }
    addButton(cb) {
      cb({
        setButtonText() { return this; },
        setCta() { return this; },
        onClick() { return this; },
      });
      return this;
    }
  }
  class Notice {
    constructor(msg) {
      this.message = msg;
    }
  }
  return {
    Modal,
    Setting,
    Notice,
    TFile: MockTFile,
    TFolder: MockTFolder,
    requestUrl: async (opts) => activeHandler(opts),
  };
}

let bundle = null;
async function loadBundle() {
  if (bundle) return bundle;
  const built = await esbuild.build({
    stdin: {
      contents: `
        export { GraphVaultClient } from "./src/graph.ts";
        export { CollaborativeSyncEngine } from "./src/sync.ts";
        export { ShareVaultModal } from "./src/shareModal.ts";
        export { EntraAuthClient } from "./src/auth.ts";
      `,
      resolveDir: process.cwd(),
      loader: "ts",
    },
    bundle: true,
    platform: "node",
    format: "cjs",
    external: ["obsidian"],
    write: false,
  });

  const code = built.outputFiles[0].text;
  const obsidianShim = createObsidianShim();
  const customRequire = (id) => {
    if (id === "obsidian") return obsidianShim;
    return Module.createRequire(import.meta.url)(id);
  };
  const mod = { exports: {} };
  const fn = new Function("exports", "require", "module", code);
  fn(mod.exports, customRequire, mod);
  bundle = mod.exports;
  return bundle;
}

function makeResponse(entry, bodyOverride) {
  const status = entry.status;
  const json = bodyOverride !== undefined ? bodyOverride : entry.body;
  const text = typeof json === "string" ? json : JSON.stringify(json);
  const u8 = Buffer.from(text, "utf8");
  return {
    status,
    headers: entry.headers || { "content-type": "application/json" },
    text,
    json: typeof json === "string" ? undefined : json,
    arrayBuffer: u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength),
  };
}

test("Sanitized Graph fixtures contain zero non-synthetic emails, tenant domains, UUIDs, drive IDs, or Base64 claims", () => {
  for (const [label, filePath] of [
    ["consumer", consumerFixturePath],
    ["enterprise", enterpriseFixturePath],
  ]) {
    const rawText = fs.readFileSync(filePath, "utf8");

    for (const m of rawText.matchAll(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g)) {
      assert.match(
        m[0],
        /^user\d+@(outlook|contoso)\.example\.com$/,
        `Unexpected email in ${label} fixture: ${m[0]}`
      );
    }
    for (const m of rawText.matchAll(/[a-zA-Z0-9-]+\.(sharepoint|onmicrosoft)\.com/gi)) {
      assert.match(
        m[0],
        /^(contoso|contoso-my)\.(sharepoint|onmicrosoft)\.com$/i,
        `Unexpected domain in ${label} fixture: ${m[0]}`
      );
    }
    for (const m of rawText.matchAll(
      /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g
    )) {
      assert.match(
        m[0],
        /^00000000-0000-4000-8000-0000000000\d\d$/,
        `Unexpected UUID in ${label} fixture: ${m[0]}`
      );
    }
    for (const m of rawText.matchAll(/\bb(?:!|%21)[A-Za-z0-9_-]{12,}/gi)) {
      assert.match(
        m[0],
        /SyntheticSharePointDriveId/i,
        `Unexpected SharePoint driveId in ${label} fixture: ${m[0]}`
      );
    }
    for (const m of rawText.matchAll(/\b01[A-Z2-7]{32}\b/g)) {
      assert.ok(
        m[0].startsWith("01SYNTHETICITEM"),
        `Unexpected SharePoint item ID in ${label} fixture: ${m[0]}`
      );
    }
    assert.equal(/<REDACTED_ACCESS_TOKEN>/.test(rawText), true);
    assert.equal(/eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}/.test(rawText), false);
  }
});

test("Consumer (OneDrive Personal) live Graph fixture contract: folder conflict fallback, personal IDs, siteUser permissions, and name-less delta tombstones", async () => {
  const { GraphVaultClient, CollaborativeSyncEngine, ShareVaultModal } = await loadBundle();
  const ep = consumerFixture.endpoints;

  const postedBodies = [];
  let deltaStage = "initial";
  let putStage = "create";
  let duplicateVaultMode = false;

  activeHandler = async (opts) => {
    const method = opts.method || "GET";
    const url = opts.url;

    if (method === "GET" && url.endsWith("/me")) {
      return makeResponse(ep["GET /v1.0/me"]);
    }
    if (method === "POST" && url.endsWith("/me/drive/root/children")) {
      postedBodies.push(JSON.parse(opts.body));
      return makeResponse(ep["POST /v1.0/me/drive/root/children (create parent folder)"]);
    }
    if (method === "POST" && /\/me\/drive\/items\/[^/]+\/children$/.test(url)) {
      const postedBody = JSON.parse(opts.body);
      postedBodies.push(postedBody);
      if (postedBody.name === ".obsidian") {
        return {
          status: 201,
          headers: {},
          json: { id: "marker-folder", name: ".obsidian", folder: {} },
          text: JSON.stringify({ id: "marker-folder", name: ".obsidian", folder: {} }),
        };
      }
      if (duplicateVaultMode) {
        return makeResponse(ep["POST /v1.0/me/drive/items/{parentId}/children (duplicate folder 409)"]);
      }
      return makeResponse(ep["POST /v1.0/me/drive/items/{parentId}/children (create vault)"]);
    }
    if (method === "GET" && url.endsWith(":/.obsidian")) {
      return {
        status: 200,
        headers: {},
        json: { id: "marker-folder", name: ".obsidian", folder: {} },
        text: JSON.stringify({ id: "marker-folder", name: ".obsidian", folder: {} }),
      };
    }
    if (method === "GET" && url.includes("/me/drive/root:/_ObsidianSyncFixtureProbe/TestVault")) {
      return {
        status: 200,
        headers: {},
        json: ep["POST /v1.0/me/drive/items/{parentId}/children (create vault)"].body,
        text: JSON.stringify(ep["POST /v1.0/me/drive/items/{parentId}/children (create vault)"].body),
      };
    }
    if (method === "GET" && url.endsWith(":/children")) {
      if (url.includes("_ObsidianSyncFixtureProbe")) {
        return makeResponse(ep["GET /v1.0/me/drive/root:/{folder}:/children"]);
      }
      return { status: 404, headers: {}, json: { error: { code: "itemNotFound" } }, text: "{}" };
    }
    if (method === "GET" && url.endsWith("/me/drive/sharedWithMe")) {
      return makeResponse(ep["GET /v1.0/me/drive/sharedWithMe"]);
    }
    if (method === "POST" && /\/drives\/[^/]+\/items\/[^/]+\/children$/.test(url)) {
      const postedBody = JSON.parse(opts.body);
      assert.equal(postedBody.name, ".obsidian");
      return {
        status: 201,
        headers: {},
        json: { id: "marker-folder", name: ".obsidian", folder: {} },
        text: JSON.stringify({ id: "marker-folder", name: ".obsidian", folder: {} }),
      };
    }
    if (method === "GET" && url.includes("/shares/")) {
      return makeResponse(ep["GET /v1.0/shares/{u!token}/driveItem (TestVault URL)"]);
    }
    if (method === "GET" && url.includes("/delta")) {
      if (deltaStage === "initial") {
        return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (initial empty)"]);
      }
      if (deltaStage === "create") {
        return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (incremental create)"]);
      }
      return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (incremental delete tombstone)"]);
    }
    if (method === "PUT" && url.endsWith(":/content")) {
      if (putStage === "create") {
        return makeResponse(ep["PUT /v1.0/drives/{driveId}/items/{vaultId}:/{path}:/content (create v1)"]);
      }
      if (putStage === "stale") {
        putStage = "update";
        return makeResponse(ep["PUT /v1.0/drives/{driveId}/items/{vaultId}:/{path}:/content (stale If-Match -> 412)"]);
      }
      return makeResponse(ep["PUT /v1.0/drives/{driveId}/items/{vaultId}:/{path}:/content (update v2 with If-Match)"]);
    }
    if (method === "GET" && url.includes("/redacted-download-url")) {
      const content = "# Fixture Probe\nLine 2\n";
      const u8 = Buffer.from(content, "utf8");
      return {
        status: 200,
        headers: { "content-type": "text/markdown" },
        text: content,
        arrayBuffer: u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength),
      };
    }
    if (method === "GET" && url.endsWith("/versions")) {
      return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{itemId}/versions"]);
    }
    if (method === "GET" && url.endsWith("/permissions")) {
      return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{vaultId}/permissions"]);
    }
    if (method === "GET" && /\/drives\/[^/]+\/items\/[^/]+$/.test(url)) {
      return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{itemId}"]);
    }
    throw new Error(`Unhandled Consumer Graph URL: ${method} ${url}`);
  };

  const graph = new GraphVaultClient(async () => "consumer-token");

  // 1. Folder creation uses @microsoft.graph.conflictBehavior = "fail" and handles 409 fallback
  const createdTarget = await graph.createOrGetVaultFolder("TestVault", "_ObsidianSyncFixtureProbe");
  assert.equal(createdTarget.driveId, "0000000000000001");
  assert.equal(createdTarget.rootItemId, "0000000000000002!s00000000000000000000000000000003");
  assert.equal(postedBodies[0]["@microsoft.graph.conflictBehavior"], "fail");

  duplicateVaultMode = true;
  const existingTarget = await graph.createOrGetVaultFolder("TestVault", "_ObsidianSyncFixtureProbe");
  assert.equal(existingTarget.rootItemId, createdTarget.rootItemId);

  // 2. Discover vaults: personal drive folder included, shared non-folder Excel file ignored
  const discovered = await graph.listAvailableVaults(false, "_ObsidianSyncFixtureProbe");
  assert.equal(discovered.length, 1);
  assert.equal(discovered[0].name, "TestVault (My OneDrive)");

  // 3. URL resolution via /shares/{u!token}/driveItem
  const resolved = await graph.resolveVaultFromUrl(
    "https://onedrive.live.com/?cid=REDACTED_CID&id=REDACTED_ID"
  );
  assert.equal(resolved.rootItemId, "0000000000000002!s00000000000000000000000000000003");

  // 4. Sync engine against real Consumer /delta (root item ignored, create pulled, name-less tombstone deleted)
  const localVault = new MockVault();
  const savedStore = { deltaLink: null, files: {}, folders: {} };
  const sync = new CollaborativeSyncEngine(
    localVault,
    graph,
    createdTarget,
    savedStore,
    async () => {}
  );

  const initRes = await sync.pullRemoteChanges();
  assert.equal(initRes.pulled, 0, "Root folder item in initial /delta must not be treated as a child file");

  deltaStage = "create";
  const createRes = await sync.pullRemoteChanges();
  assert.equal(createRes.pulled, 1);
  assert.ok(localVault.getAbstractFileByPath("Notes/Hello.md"), "Notes/Hello.md should be created locally");

  // Stale ETag 412 triggers 3-way merge using real Graph 412 error shape
  await localVault.modify(
    localVault.getAbstractFileByPath("Notes/Hello.md"),
    "# Fixture Probe\nLine 2\nLocal addition\n"
  );
  putStage = "stale";
  const mergePush = await sync.pushPath("Notes/Hello.md");
  assert.equal(mergePush.merged, 1);

  // Real Consumer /delta tombstone has NO `name`, NO `eTag`, and NO `parentReference.path`!
  const tombstoneEntry =
    ep["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (incremental delete tombstone)"].body.value[1];
  assert.equal(tombstoneEntry.name, undefined);
  assert.equal(tombstoneEntry.parentReference.path, undefined);

  deltaStage = "delete";
  const delRes = await sync.pullRemoteChanges();
  assert.equal(delRes.pulled, 1, "Tracked file should be deleted via itemId even when tombstone omits name");
  assert.equal(localVault.getAbstractFileByPath("Notes/Hello.md"), null);

  // Repeating the same name-less tombstone after deletion must be a clean no-op (not "undefined")
  const repeatDelRes = await sync.pullRemoteChanges();
  assert.equal(repeatDelRes.pulled, 0);
  assert.equal(localVault.files.size, 0);

  // 5. Consumer OneDrive /permissions uses `grantedToV2.siteUser` & `grantedTo.user` (no `grantedToV2.user`)
  const modal = new ShareVaultModal({}, createdTarget, graph);
  await modal.onOpen();
  const renderedText = modal.contentEl.findAllText().join("\n");
  assert.match(
    renderedText,
    /Alex Rivera \(user2@outlook\.example\.com\) — \[owner\]/,
    "ShareVaultModal must render Consumer OneDrive siteUser/grantedTo displayName + email"
  );
});

test("Enterprise (Entra ID / SharePoint / Teams) live Graph fixture contract: localized Dutch library URLs, SharePoint discovery, and deep nested name-less tombstones", async () => {
  const { GraphVaultClient, CollaborativeSyncEngine, ShareVaultModal } = await loadBundle();
  const ep = enterpriseFixture.endpoints;

  let sharesReturn403 = false;
  let deltaPayloadOverride = null;

  activeHandler = async (opts) => {
    const method = opts.method || "GET";
    const url = opts.url;

    if (method === "GET" && url.endsWith(":/children") && url.includes("/me/drive/")) {
      if (url.includes("_ObsidianSyncFixtureProbe")) {
        return makeResponse(ep["GET /v1.0/me/drive/root:/{folder}:/children"]);
      }
      return { status: 404, headers: {}, json: { error: { code: "itemNotFound" } }, text: "{}" };
    }
    if (method === "GET" && url.endsWith("/me/drive/sharedWithMe")) {
      return makeResponse(ep["GET /v1.0/me/drive/sharedWithMe"]);
    }
    if (method === "GET" && url.endsWith("/me/followedSites")) {
      return makeResponse(ep["GET /v1.0/me/followedSites"]);
    }
    if (method === "GET" && /\/sites\/[^/:]+\/drive$/.test(url)) {
      return makeResponse(ep["GET /v1.0/sites/{siteId}/drive"]);
    }
    if (method === "GET" && /\/sites\/[^/:]+\/drives$/.test(url)) {
      return {
        status: 200,
        headers: {},
        json: { value: [ep["GET /v1.0/sites/{siteId}/drive"].body] },
        text: JSON.stringify({ value: [ep["GET /v1.0/sites/{siteId}/drive"].body] }),
      };
    }
    if (method === "GET" && /\/sites\/contoso\.sharepoint\.com:\/sites\/SampleSite_11$/i.test(url)) {
      return {
        status: 200,
        headers: {},
        json: {
          id: "contoso.sharepoint.com,00000000-0000-4000-8000-000000000019,00000000-0000-4000-8000-000000000020",
          displayName: "SampleSite_11",
          webUrl: "https://contoso.sharepoint.com/sites/SampleSite_11",
        },
        text: "{}",
      };
    }
    if (method === "GET" && /\/drives\/[^/]+\/root\/children$/.test(url)) {
      return makeResponse(ep["GET /v1.0/drives/{spDriveId}/root/children"]);
    }
    if (method === "GET" && /\/drives\/[^/]+\/root:\/TeamVault$/.test(url)) {
      return {
        status: 200,
        headers: {},
        json: {
          id: "01SYNTHETICITEM0000000000000000099",
          name: "TeamVault",
          folder: { childCount: 0 },
          webUrl: "https://contoso.sharepoint.com/sites/SampleSite_11/Gedeelde%20Documenten/TeamVault",
        },
        text: "{}",
      };
    }
    if (method === "GET" && url.includes("/shares/")) {
      if (sharesReturn403) {
        return { status: 403, headers: {}, json: { error: { code: "accessDenied" } }, text: "{}" };
      }
      return makeResponse(ep["GET /v1.0/shares/{u!token}/driveItem (SharePoint/Teams URL)"]);
    }
    if (method === "GET" && url.includes("/delta")) {
      if (deltaPayloadOverride) {
        return { status: 200, headers: {}, json: deltaPayloadOverride, text: JSON.stringify(deltaPayloadOverride) };
      }
      return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{vaultId}/delta (incremental create)"]);
    }
    if (method === "PUT" && url.endsWith(":/content")) {
      return makeResponse(ep["PUT /v1.0/drives/{driveId}/items/{vaultId}:/{path}:/content (create v1)"]);
    }
    if (method === "GET" && url.includes("/redacted-download-url")) {
      const content = "# Enterprise Fixture\n";
      const u8 = Buffer.from(content, "utf8");
      return {
        status: 200,
        headers: { "content-type": "text/markdown" },
        text: content,
        arrayBuffer: u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength),
      };
    }
    if (method === "GET" && url.endsWith(":/.obsidian")) {
      return {
        status: 200,
        headers: {},
        json: { id: "marker-folder", name: ".obsidian", folder: {} },
        text: JSON.stringify({ id: "marker-folder", name: ".obsidian", folder: {} }),
      };
    }
    if (method === "GET" && url.endsWith("/permissions")) {
      return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{vaultId}/permissions"]);
    }
    if (method === "GET" && /\/drives\/[^/]+\/items\/[^/]+$/.test(url)) {
      return makeResponse(ep["GET /v1.0/drives/{driveId}/items/{itemId}"]);
    }
    throw new Error(`Unhandled Enterprise Graph URL: ${method} ${url}`);
  };

  const graph = new GraphVaultClient(async () => "enterprise-token");

  // 1. Enterprise discovery combines OneDrive for Business + followed SharePoint libraries + SharePoint subfolders
  const vaults = await graph.listAvailableVaults(true, "_ObsidianSyncFixtureProbe");
  assert.ok(vaults.some((v) => v.name === "TestVault (My OneDrive)"));
  assert.ok(vaults.some((v) => v.name === "SampleSite_5 (SharePoint)" && v.rootItemId === "root"));
  assert.ok(
    vaults.some(
      (v) =>
        v.name === "SampleFolder_17 (SharePoint: SampleSite_5)" &&
        v.rootItemId === "01SYNTHETICITEM0000000000000000008"
    )
  );

  // 2. Direct SharePoint/Teams URL resolution via /shares/{u!token}/driveItem
  const spDirect = await graph.resolveVaultFromUrl(
    "https://contoso.sharepoint.com/sites/SampleSite_21/Gedeelde%20Documenten/SampleFolder_20"
  );
  assert.equal(spDirect.driveId, "b!SyntheticSharePointDriveId000500000000000000000000");
  assert.equal(spDirect.rootItemId, "01SYNTHETICITEM0000000000000000010");

  // 3. Fallback resolution on localized Dutch SharePoint URL ("Gedeelde Documenten" where drive.name !== "Gedeelde Documenten")
  sharesReturn403 = true;
  const spDutchFallback = await graph.resolveVaultFromUrl(
    "https://contoso.sharepoint.com/sites/SampleSite_11/Gedeelde%20Documenten/TeamVault"
  );
  assert.equal(spDutchFallback.driveId, "b!SyntheticSharePointDriveId000400000000000000000000");
  assert.equal(spDutchFallback.rootItemId, "01SYNTHETICITEM0000000000000000099");
  assert.equal(spDutchFallback.name, "TeamVault (SharePoint: SampleSite_11)");

  // 4. Sync engine against real Enterprise /delta + deep nested name-less folder tombstones
  const target = {
    name: "TestVault (My OneDrive)",
    driveId: "b!SyntheticSharePointDriveId000100000000000000000000",
    rootItemId: "01SYNTHETICITEM0000000000000000003",
    webUrl: "",
  };
  const localVault = new MockVault();
  const savedStore = { deltaLink: null, files: {}, folders: {} };
  const sync = new CollaborativeSyncEngine(
    localVault,
    graph,
    target,
    savedStore,
    async () => {}
  );

  const pulled = await sync.pullRemoteChanges();
  assert.equal(pulled.pulled, 1);
  assert.ok(localVault.getAbstractFileByPath("Notes/Hello.md"));

  // Deep nested folder tree deleted remotely where ALL tombstones omit `name` and `parentReference.path`
  // (ordered parent-first, matching SharePoint's worst-case ordering)
  await localVault.createFolder("Projects");
  await localVault.createFolder("Projects/Q1");
  await localVault.create("Projects/Q1/Spec.md", "# Spec\n");
  savedStore.files["Projects/Q1/Spec.md"] = {
    itemId: "01SYNTHETICITEM_SPEC",
    eTag: "\"{00000000-0000-4000-8000-000000000090},1\"",
    baseText: "# Spec\n",
    lastSyncedMtime: Date.now(),
  };
  // Intentionally do NOT pre-populate savedStore.folders["Projects"] or ["Projects/Q1"]
  // to verify parentReference.id upward inference from tombstones alone!
  deltaPayloadOverride = {
    "@odata.deltaLink": "https://graph.microsoft.com/v1.0/drives/b!Synthetic/items/root/delta?token=2",
    value: [
      {
        id: "01SYNTHETICITEM_PROJECTS",
        deleted: { state: "deleted" },
        folder: { childCount: 0 },
        parentReference: {
          driveId: target.driveId,
          id: target.rootItemId,
        },
      },
      {
        id: "01SYNTHETICITEM_Q1",
        deleted: { state: "deleted" },
        folder: { childCount: 0 },
        parentReference: {
          driveId: target.driveId,
          id: "01SYNTHETICITEM_PROJECTS",
        },
      },
      {
        id: "01SYNTHETICITEM_SPEC",
        deleted: { state: "deleted" },
        file: { hashes: { quickXorHash: "AAAAAAAAAAAAAAAAAAAAAAAAAAA=" } },
        parentReference: {
          driveId: target.driveId,
          id: "01SYNTHETICITEM_Q1",
        },
      },
    ],
  };

  await sync.pullRemoteChanges();
  assert.equal(localVault.getAbstractFileByPath("Projects/Q1/Spec.md"), null);
  assert.equal(localVault.getAbstractFileByPath("Projects/Q1"), null);
  assert.equal(localVault.getAbstractFileByPath("Projects"), null);

  // 5. Enterprise /permissions renders Jordan Patel
  const modal = new ShareVaultModal({}, target, graph);
  await modal.onOpen();
  const renderedText = modal.contentEl.findAllText().join("\n");
  assert.match(renderedText, /Jordan Patel \(user2@contoso\.example\.com\) — \[owner\]/);
});
