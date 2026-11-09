import test from "node:test";
import assert from "node:assert/strict";
import Module from "node:module";
import { createMockGraphServer } from "./mock-graph-server.mjs";

// Provide a lightweight mock of the "obsidian" runtime module so we can import the compiled main.js bundle directly in Node
class MockTFile {
  constructor(path, content = "") {
    this.path = path;
    this.name = path.split("/").pop();
    this.extension = this.name.includes(".") ? this.name.split(".").pop() : "";
    this.stat = { mtime: Date.now(), ctime: Date.now(), size: content.length };
  }
}

class MockTFolder {
  constructor(path) {
    this.path = path;
    this.name = path.split("/").pop();
  }
}

class MockVault {
  constructor() {
    this.configDir = ".obsidian";
    this.files = new Map();
    this.folders = new Map();
    this.adapter = {
      exists: async (p) => this.files.has(p) || this.folders.has(p),
      list: async (dir) => {
        const clean = dir.replace(/^\/+|\/+$/g, "");
        const prefix = clean ? `${clean}/` : "";
        const files = [];
        const folders = [];
        for (const k of this.files.keys()) {
          if (k.startsWith(prefix)) {
            const rest = k.slice(prefix.length);
            if (rest && !rest.includes("/")) {
              files.push(k);
            }
          }
        }
        for (const k of this.folders.keys()) {
          if (k.startsWith(prefix)) {
            const rest = k.slice(prefix.length);
            if (rest && !rest.includes("/")) {
              folders.push(k);
            }
          }
        }
        return { files, folders };
      },
      read: async (p) => this.files.get(p)?.content ?? "",
      readBinary: async (p) => {
        const entry = this.files.get(p);
        if (entry?.binary) return entry.binary;
        const u8 = Buffer.from(entry?.content ?? "", "utf8");
        return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
      },
      write: async (p, content) => {
        await this.create(p, content);
      },
      writeBinary: async (p, arrayBuffer) => {
        await this.createBinary(p, arrayBuffer);
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
    // Just like real Obsidian, vault.getFiles() does NOT return .obsidian/ config files
    return Array.from(this.files.values())
      .map((f) => f.file)
      .filter((f) => !f.path.startsWith(".obsidian/"));
  }
  getMarkdownFiles() {
    return Array.from(this.files.values()).filter((f) => f.file.extension === "md").map((f) => f.file);
  }
  getAbstractFileByPath(p) {
    if (this.files.has(p)) return this.files.get(p).file;
    if (this.folders.has(p)) return this.folders.get(p);
    return null;
  }
  async read(file) {
    return this.files.get(file.path)?.content ?? "";
  }
  async readBinary(file) {
    const entry = this.files.get(file.path);
    if (entry?.binary) return entry.binary;
    const text = entry?.content ?? "";
    const u8 = Buffer.from(text, "utf8");
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
  }
  async create(path, content) {
    const file = new MockTFile(path, content);
    this.files.set(path, { file, content });
    return file;
  }
  async createBinary(path, arrayBuffer) {
    const buf = Buffer.from(arrayBuffer);
    const file = new MockTFile(path, "");
    file.stat.size = buf.byteLength;
    this.files.set(path, {
      file,
      content: "",
      binary: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
    });
    return file;
  }
  async modify(file, content) {
    const entry = this.files.get(file.path);
    if (entry) {
      entry.content = content;
      entry.file.stat.mtime = Date.now();
      entry.file.stat.size = content.length;
    } else {
      await this.create(file.path, content);
    }
  }
  async modifyBinary(file, arrayBuffer) {
    const buf = Buffer.from(arrayBuffer);
    const entry = this.files.get(file.path);
    if (entry) {
      entry.binary = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
      entry.file.stat.mtime = Date.now();
      entry.file.stat.size = buf.byteLength;
    } else {
      await this.createBinary(file.path, arrayBuffer);
    }
  }
  async delete(fileOrFolder) {
    this.files.delete(fileOrFolder.path);
    this.folders.delete(fileOrFolder.path);
    const prefix = `${fileOrFolder.path}/`;
    for (const k of Array.from(this.files.keys())) {
      if (k.startsWith(prefix)) this.files.delete(k);
    }
    for (const k of Array.from(this.folders.keys())) {
      if (k.startsWith(prefix)) this.folders.delete(k);
    }
  }
  async createFolder(path) {
    const folder = new MockTFolder(path);
    this.folders.set(path, folder);
    return folder;
  }
}

const obsidianMock = {
  TFile: MockTFile,
  TFolder: MockTFolder,
  Vault: MockVault,
  Plugin: class {},
  PluginSettingTab: class {},
  Modal: class {},
  Setting: class {},
  Notice: class {},
  requestUrl: async ({ url, method = "GET", headers = {}, body, throw: shouldThrow = true }) => {
    const res = await fetch(url, {
      method,
      headers,
      body: body instanceof ArrayBuffer ? Buffer.from(body) : body,
    });
    const arrayBuffer = await res.arrayBuffer();
    const text = Buffer.from(arrayBuffer).toString("utf8");
    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      // non-json response
    }
    if (!res.ok && shouldThrow !== false) {
      throw new Error(`HTTP ${res.status}: ${text}`);
    }
    return {
      status: res.status,
      text,
      json,
      arrayBuffer,
    };
  },
};

const origLoad = Module._load;
Module._load = function (request, parent, isMain) {
  if (request === "obsidian") {
    return obsidianMock;
  }
  return origLoad.apply(this, arguments);
};

test("End-to-end Microsoft Entra ID Auth + SharePoint/OneDrive Graph Vault Sharing, 3-Way Merge, and Version History", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const baseUrl = `http://127.0.0.1:${port}/v1.0`;

  try {
    // Load the esbuild-compiled plugin bundle or source modules via esbuild register
    const esbuild = await import("esbuild");
    const authBuild = await esbuild.build({
      entryPoints: ["src/auth.ts", "src/graph.ts", "src/sync.ts", "src/logger.ts"],
      bundle: true,
      format: "cjs",
      platform: "node",
      external: ["obsidian"],
      write: false,
      outdir: "out",
    });

    const loadBundled = (name) => {
      const file = authBuild.outputFiles.find((f) => f.path.endsWith(name));
      const m = { exports: {} };
      const fn = new Function("exports", "require", "module", file.text);
      fn(m.exports, (id) => (id === "obsidian" ? obsidianMock : origLoad(id)), m);
      return m.exports;
    };

    const { EntraPkceAuth } = loadBundled("auth.js");
    const { GraphVaultClient } = loadBundled("graph.js");
    const { CollaborativeSyncEngine } = loadBundled("sync.js");
    const { SyncLogger } = loadBundled("logger.js");

    const clientId = process.env.ENTRA_CLIENT_ID || "ada87c12-a404-484b-a7dc-10c54beb2d7d";
    const pass = process.env.ENTRA_TEST_PASSWORD;

    let aliceTokens;
    let bobTokens;
    if (pass) {
      // 1. Authenticate both real Entra ID users (Alice & Bob) against live Microsoft Entra ID
      const aliceAuth = new EntraPkceAuth({
        tenantId: "common",
        clientId,
        includeSitesScope: false,
      });
      const bobAuth = new EntraPkceAuth({
        tenantId: "common",
        clientId,
        includeSitesScope: false,
      });

      aliceTokens = await aliceAuth.signInWithCredentials(
        "alice@woutervanranstgmail.onmicrosoft.com",
        pass
      );
      bobTokens = await bobAuth.signInWithCredentials(
        "bob@woutervanranstgmail.onmicrosoft.com",
        pass
      );
    } else {
      // Fallback unsigned test JWTs when running in CI without ENTRA_TEST_PASSWORD
      const makeJwt = (oid, name, upn) => {
        const header = Buffer.from(JSON.stringify({ alg: "none", typ: "JWT" })).toString("base64url");
        const payload = Buffer.from(JSON.stringify({ oid, name, upn })).toString("base64url");
        return `${header}.${payload}.`;
      };
      aliceTokens = {
        accessToken: makeJwt(
          "63a6c51d-b9c7-43cf-8ba0-ce82bb6eb2e6",
          "Alice Peeters",
          "alice@woutervanranstgmail.onmicrosoft.com"
        ),
        refreshToken: "refresh-alice",
        expiresAt: Date.now() + 3600_000,
        userDisplayName: "Alice Peeters",
        userEmail: "alice@woutervanranstgmail.onmicrosoft.com",
      };
      bobTokens = {
        accessToken: makeJwt(
          "3be7d8b4-7e90-4da0-9499-619b3da4e2cf",
          "Bob Janssens",
          "bob@woutervanranstgmail.onmicrosoft.com"
        ),
        refreshToken: "refresh-bob",
        expiresAt: Date.now() + 3600_000,
        userDisplayName: "Bob Janssens",
        userEmail: "bob@woutervanranstgmail.onmicrosoft.com",
      };
    }

    assert.equal(aliceTokens.userDisplayName, "Alice Peeters");
    assert.equal(aliceTokens.userEmail, "alice@woutervanranstgmail.onmicrosoft.com");
    assert.equal(bobTokens.userDisplayName, "Bob Janssens");
    assert.equal(bobTokens.userEmail, "bob@woutervanranstgmail.onmicrosoft.com");

    const aliceGraph = new GraphVaultClient(
      async () => aliceTokens.accessToken,
      () => baseUrl
    );
    const bobGraph = new GraphVaultClient(
      async () => bobTokens.accessToken,
      () => baseUrl
    );

    // 2. Alice creates a shared vault in "/Documents/Documents/EnterpriseArchitecture" (matching consumer OneDrive + SharePoint pattern)
    const aliceVaultTarget = await aliceGraph.createOrGetVaultFolder(
      "EnterpriseArchitecture",
      "Documents/Documents"
    );
    assert.ok(aliceVaultTarget.driveId);
    assert.ok(aliceVaultTarget.rootItemId);

    // 3. Alice shares the vault with Bob via per-vault /invite
    const invitePerms = await aliceGraph.inviteCollaborators(
      aliceVaultTarget,
      ["bob@woutervanranstgmail.onmicrosoft.com"],
      "write"
    );
    assert.equal(invitePerms.length, 1);

    // 4. Bob discovers the shared vault in his /me/drive/sharedWithMe
    const bobDiscovered = await bobGraph.listAvailableVaults(false, "Documents/Documents");
    assert.equal(bobDiscovered.length, 1);
    assert.equal(bobDiscovered[0].driveId, aliceVaultTarget.driveId);
    assert.equal(bobDiscovered[0].rootItemId, aliceVaultTarget.rootItemId);
    const bobVaultTarget = bobDiscovered[0];

    // 5. Set up Alice and Bob's local Obsidian vaults and sync engines
    const aliceObsidianVault = new MockVault();
    const bobObsidianVault = new MockVault();
    const aliceStore = { files: {} };
    const bobStore = { files: {} };

    const aliceSync = new CollaborativeSyncEngine(
      aliceObsidianVault,
      aliceGraph,
      aliceVaultTarget,
      aliceStore,
      async () => {}
    );
    const bobSync = new CollaborativeSyncEngine(
      bobObsidianVault,
      bobGraph,
      bobVaultTarget,
      bobStore,
      async () => {}
    );

    // 6. Alice creates Roadmap.md and syncs it to SharePoint/OneDrive
    const initialNote = [
      "# Enterprise Roadmap 2026",
      "",
      "## Section A: Identity",
      "- Owner: Alice",
      "- Status: Draft",
      "",
      "## Section B: Storage",
      "- Backend: SharePoint Online",
      "- Status: Pending",
    ].join("\n");

    await aliceObsidianVault.create("Roadmap.md", initialNote);
    const aliceInitialSync = await aliceSync.syncAll();
    assert.equal(aliceInitialSync.pushed, 1);

    // 7. Bob runs syncAll() and receives Roadmap.md via Graph /delta
    const bobInitialSync = await bobSync.syncAll();
    assert.equal(bobInitialSync.pulled, 1);
    const bobFile = bobObsidianVault.getAbstractFileByPath("Roadmap.md");
    assert.ok(bobFile);
    assert.equal(await bobObsidianVault.read(bobFile), initialNote);

    // 8. Concurrent edits!
    // Alice updates Section A and pushes to SharePoint first (creating Version 2.0)
    const aliceEdit = initialNote.replace("- Status: Draft", "- Status: Approved by Alice");
    const aliceFile = aliceObsidianVault.getAbstractFileByPath("Roadmap.md");
    await aliceObsidianVault.modify(aliceFile, aliceEdit);
    await aliceSync.pushMarkdownFile(aliceFile);

    // Meanwhile, Bob (still on Version 1.0 eTag!) edits Section B locally and pushes.
    // The server returns HTTP 412 Precondition Failed, and Bob's sync engine automatically runs 3-way diff3 merge!
    const bobEdit = initialNote.replace(
      "- Status: Pending",
      "- Status: Configured by Bob"
    );
    await bobObsidianVault.modify(bobFile, bobEdit);
    const bobPushResult = await bobSync.pushMarkdownFile(bobFile);
    assert.equal(bobPushResult.merged, 1);
    assert.equal(bobPushResult.conflicts, 0);

    const mergedContent = await bobObsidianVault.read(bobFile);
    assert.ok(mergedContent.includes("- Status: Approved by Alice"), "Preserved Alice's edit");
    assert.ok(mergedContent.includes("- Status: Configured by Bob"), "Preserved Bob's edit");

    // 9. Alice pulls the merged version
    await aliceSync.syncAll();
    assert.equal(await aliceObsidianVault.read(aliceFile), mergedContent);

    // 10. Check Version History with Author Attribution!
    const itemId = aliceStore.files["Roadmap.md"].itemId;
    const versions = await aliceGraph.listVersions(aliceVaultTarget.driveId, itemId);
    assert.equal(versions.length, 3);
    assert.equal(versions[0].id, "3.0");
    assert.equal(versions[0].lastModifiedBy.user.displayName, "Bob Janssens");
    assert.equal(versions[1].id, "2.0");
    assert.equal(versions[1].lastModifiedBy.user.displayName, "Alice Peeters");
    assert.equal(versions[2].id, "1.0");
    assert.equal(versions[2].lastModifiedBy.user.displayName, "Alice Peeters");

    // Verify restoring Version 1.0 content
    const v1Text = await aliceGraph.downloadVersionText(
      aliceVaultTarget.driveId,
      itemId,
      "1.0"
    );
    assert.equal(v1Text, initialNote);

    // 11. Verify HTTP 401 automatic token refresh retry + SyncLogger diagnostics + downloadUrl 401 fallback
    const testLogger = new SyncLogger();
    let forceRefreshCalled = false;
    const refreshingGraph = new GraphVaultClient(
      async (forceRefresh = false) => {
        if (forceRefresh) {
          forceRefreshCalled = true;
          return aliceTokens.accessToken;
        }
        // Return an invalid/expired token on first attempt so the server responds with 401
        return "expired-non-jwt-token";
      },
      () => baseUrl,
      testLogger
    );

    const deltaWithAutoRefresh = await refreshingGraph.fetchDelta(aliceVaultTarget);
    assert.equal(forceRefreshCalled, true, "Should automatically force token refresh after 401");
    assert.ok(deltaWithAutoRefresh.items.length > 0);
    assert.ok(
      testLogger.formatText().includes("returned 401 Unauthorized"),
      "SyncLogger should record the 401 warning and token refresh retry"
    );

    // Verify downloadFileText fallback when @microsoft.graph.downloadUrl returns 401
    const fallbackResult = await aliceGraph.downloadFileText(
      aliceVaultTarget.driveId,
      itemId
    );
    assert.equal(fallbackResult.text, mergedContent);

    // 12. Verify Binary Attachment Sync (.png image inside subfolder)
    await aliceObsidianVault.createFolder("assets");
    const pngBytes = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0xff, 0xfe, 0xfd, 0x42,
    ]);
    const pngArrayBuffer = pngBytes.buffer.slice(
      pngBytes.byteOffset,
      pngBytes.byteOffset + pngBytes.byteLength
    );
    await aliceObsidianVault.createBinary("assets/architecture.png", pngArrayBuffer);

    const aliceAttachmentSync = await aliceSync.syncAll();
    assert.equal(aliceAttachmentSync.pushed, 1, "Alice should push binary attachment");

    const bobAttachmentSync = await bobSync.syncAll();
    assert.equal(bobAttachmentSync.pulled, 1, "Bob should pull binary attachment");
    const bobPngFile = bobObsidianVault.getAbstractFileByPath("assets/architecture.png");
    assert.ok(bobPngFile, "Bob vault should have assets/architecture.png");
    const bobPngBytes = new Uint8Array(await bobObsidianVault.readBinary(bobPngFile));
    assert.deepEqual(
      Array.from(bobPngBytes),
      Array.from(pngBytes),
      "Binary attachment bytes must match identically"
    );

    // 13. Verify Bidirectional Page (File) Deletion: Alice deletes Roadmap.md locally -> deleted on Bob's vault
    await aliceObsidianVault.delete(aliceFile);
    const aliceDeleteSync = await aliceSync.syncAll();
    assert.equal(aliceDeleteSync.deleted, 1, "Alice should delete Roadmap.md remotely");

    const bobAfterFileDelete = await bobSync.syncAll();
    assert.ok(bobAfterFileDelete.pulled >= 1);
    assert.equal(
      bobObsidianVault.getAbstractFileByPath("Roadmap.md"),
      null,
      "Roadmap.md should be deleted from Bob's vault"
    );

    // 14. Verify Bidirectional Folder Deletion: Bob deletes folder 'assets' -> deleted on Alice's vault
    const bobAssetsFolder = bobObsidianVault.getAbstractFileByPath("assets");
    assert.ok(bobAssetsFolder);
    await bobObsidianVault.delete(bobAssetsFolder);
    await bobSync.deleteLocalPath("assets", true);

    await aliceSync.syncAll();
    assert.equal(
      aliceObsidianVault.getAbstractFileByPath("assets/architecture.png"),
      null,
      "assets/architecture.png should be deleted from Alice's vault"
    );
    assert.equal(
      aliceObsidianVault.getAbstractFileByPath("assets"),
      null,
      "assets folder should be deleted from Alice's vault"
    );

    // 15. Verify Vault-Specific Settings (.obsidian/) Sync & Security Exclusions
    await aliceObsidianVault.adapter.mkdir(".obsidian");
    await aliceObsidianVault.adapter.mkdir(".obsidian/snippets");
    await aliceObsidianVault.adapter.mkdir(".obsidian/plugins");
    await aliceObsidianVault.adapter.mkdir(".obsidian/plugins/entra-sharepoint-sync");
    await aliceObsidianVault.adapter.write(
      ".obsidian/app.json",
      JSON.stringify({ strictLineBreaks: true, showLineNumber: true }, null, 2)
    );
    await aliceObsidianVault.adapter.write(
      ".obsidian/snippets/enterprise.css",
      "body { --font-text-size: 16px; }"
    );
    // Personal files that MUST NEVER be synced across users:
    await aliceObsidianVault.adapter.write(
      ".obsidian/workspace.json",
      JSON.stringify({ main: "alice-personal-tabs" })
    );
    await aliceObsidianVault.adapter.write(
      ".obsidian/plugins/entra-sharepoint-sync/data.json",
      JSON.stringify({ tokens: { accessToken: "SECRET-ALICE-TOKEN" } })
    );

    const aliceConfigSync = await aliceSync.syncAll();
    assert.equal(
      aliceConfigSync.pushed,
      2,
      "Should push .obsidian/app.json and .obsidian/snippets/enterprise.css, excluding workspace.json and plugin data.json"
    );

    const bobConfigSync = await bobSync.syncAll();
    assert.ok(bobConfigSync.pulled >= 2);
    assert.equal(
      await bobObsidianVault.adapter.read(".obsidian/app.json"),
      JSON.stringify({ strictLineBreaks: true, showLineNumber: true }, null, 2)
    );
    assert.equal(
      await bobObsidianVault.adapter.read(".obsidian/snippets/enterprise.css"),
      "body { --font-text-size: 16px; }"
    );
    assert.equal(
      await bobObsidianVault.adapter.exists(".obsidian/workspace.json"),
      false,
      "Personal workspace.json must never sync to other users"
    );
    assert.equal(
      await bobObsidianVault.adapter.exists(".obsidian/plugins/entra-sharepoint-sync/data.json"),
      false,
      "Personal Entra ID tokens in entra-sharepoint-sync/data.json must never sync to other users"
    );
  } finally {
    await mockGraph.stop();
  }
});
