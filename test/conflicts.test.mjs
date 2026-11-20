import test from "node:test";
import assert from "node:assert/strict";
import Module from "node:module";
import esbuild from "esbuild";
import { createMockGraphServer } from "./mock-graph-server.mjs";

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
      list: async () => ({ files: [], folders: [] }),
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
    return Array.from(this.files.values())
      .map((f) => f.file)
      .filter((f) => !f.path.startsWith(".obsidian/"));
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
  async read(file) {
    return this.files.get(file.path)?.content ?? "";
  }
  async readBinary(file) {
    const entry = this.files.get(file.path);
    if (entry?.binary) return entry.binary;
    const u8 = Buffer.from(entry?.content ?? "", "utf8");
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
  }
  async create(path, content) {
    const parts = path.split("/");
    parts.pop();
    if (parts.length > 0) {
      let cur = "";
      for (const seg of parts) {
        cur = cur ? `${cur}/${seg}` : seg;
        if (!this.folders.has(cur)) {
          this.folders.set(cur, new MockTFolder(cur));
        }
      }
    }
    const existing = this.files.get(path);
    if (existing) {
      existing.content = content;
      existing.file.stat.size = content.length;
      return existing.file;
    }
    const file = new MockTFile(path, content);
    this.files.set(path, { file, content });
    return file;
  }
  async createBinary(path, arrayBuffer) {
    const parts = path.split("/");
    parts.pop();
    if (parts.length > 0) {
      let cur = "";
      for (const seg of parts) {
        cur = cur ? `${cur}/${seg}` : seg;
        if (!this.folders.has(cur)) {
          this.folders.set(cur, new MockTFolder(cur));
        }
      }
    }
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
    await this.create(file.path, content);
  }
  async modifyBinary(file, arrayBuffer) {
    await this.createBinary(file.path, arrayBuffer);
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
    let cur = "";
    let last = null;
    for (const seg of path.split("/").filter(Boolean)) {
      cur = cur ? `${cur}/${seg}` : seg;
      if (!this.folders.has(cur)) {
        this.folders.set(cur, new MockTFolder(cur));
      }
      last = this.folders.get(cur);
    }
    return last;
  }
  async renameFile(oldPath, newPath) {
    const entry = this.files.get(oldPath);
    assert.ok(entry, `Expected local file ${oldPath} to exist for rename`);
    this.files.delete(oldPath);
    if (entry.binary) {
      return await this.createBinary(newPath, entry.binary);
    }
    return await this.create(newPath, entry.content);
  }
  async renameFolder(oldFolder, newFolder) {
    const oldPrefix = `${oldFolder}/`;
    const movedFiles = [];
    for (const [k, v] of Array.from(this.files.entries())) {
      if (k.startsWith(oldPrefix)) {
        const rel = k.slice(oldPrefix.length);
        this.files.delete(k);
        movedFiles.push({ newPath: `${newFolder}/${rel}`, entry: v });
      }
    }
    for (const k of Array.from(this.folders.keys())) {
      if (k === oldFolder || k.startsWith(oldPrefix)) {
        this.folders.delete(k);
      }
    }
    const folderObj = await this.createFolder(newFolder);
    for (const m of movedFiles) {
      if (m.entry.binary) {
        await this.createBinary(m.newPath, m.entry.binary);
      } else {
        await this.create(m.newPath, m.entry.content);
      }
    }
    return folderObj;
  }
}

const obsidianMock = {
  TFile: MockTFile,
  TFolder: MockTFolder,
  Vault: MockVault,
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
      // non-JSON body
    }
    if (!res.ok && shouldThrow !== false) {
      throw new Error(`HTTP ${res.status}: ${text}`);
    }
    return { status: res.status, text, json, arrayBuffer };
  },
};

const origLoad = Module._load;
Module._load = function (request) {
  if (request === "obsidian") return obsidianMock;
  return origLoad.apply(this, arguments);
};

let cachedExports = null;
async function loadSyncModules() {
  if (cachedExports) return cachedExports;
  const built = await esbuild.build({
    entryPoints: ["src/graph.ts", "src/sync.ts"],
    bundle: true,
    format: "cjs",
    platform: "node",
    external: ["obsidian"],
    write: false,
    outdir: "out",
  });
  const load = (name) => {
    const file = built.outputFiles.find((f) => f.path.endsWith(name));
    const m = { exports: {} };
    const fn = new Function("exports", "require", "module", file.text);
    fn(m.exports, (id) => (id === "obsidian" ? obsidianMock : origLoad(id)), m);
    return m.exports;
  };
  cachedExports = {
    ...load("graph.js"),
    ...load("sync.js"),
  };
  return cachedExports;
}

function makeJwt(oid, name, upn) {
  const header = Buffer.from(JSON.stringify({ alg: "none", typ: "JWT" })).toString("base64url");
  const payload = Buffer.from(JSON.stringify({ oid, name, upn })).toString("base64url");
  return `${header}.${payload}.`;
}

async function setupTwoCollaborators() {
  const { GraphVaultClient, CollaborativeSyncEngine } = await loadSyncModules();
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const baseUrl = `http://127.0.0.1:${port}/v1.0`;

  const aliceToken = makeJwt("00000000-0000-4000-8000-000000000101", "Alice Rivera", "alice@contoso.example.com");
  const bobToken = makeJwt("00000000-0000-4000-8000-000000000102", "Bob Martens", "bob@contoso.example.com");

  const aliceGraph = new GraphVaultClient(async () => aliceToken, () => baseUrl);
  const bobGraph = new GraphVaultClient(async () => bobToken, () => baseUrl);

  const target = await aliceGraph.createOrGetVaultFolder("SharedTeamVault", "ObsidianVaults");
  await aliceGraph.inviteCollaborators(target, ["bob@contoso.example.com"], "write");

  const aliceVault = new MockVault();
  const bobVault = new MockVault();
  const aliceStore = { files: {}, folders: {} };
  const bobStore = { files: {}, folders: {} };

  const aliceSync = new CollaborativeSyncEngine(
    aliceVault,
    aliceGraph,
    target,
    aliceStore,
    async () => {}
  );
  const bobSync = new CollaborativeSyncEngine(
    bobVault,
    bobGraph,
    target,
    bobStore,
    async () => {}
  );

  return {
    mockGraph,
    target,
    aliceVault,
    bobVault,
    aliceStore,
    bobStore,
    aliceSync,
    bobSync,
    aliceGraph,
    bobGraph,
  };
}

function toBytes(arr) {
  const u8 = new Uint8Array(arr);
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

test("Conflict 1: Overlapping same-line text edits on Push (412) and Pull (/delta) emit conflict markers and resolve cleanly", async () => {
  const ctx = await setupTwoCollaborators();
  try {
    const { aliceVault, bobVault, aliceSync, bobSync } = ctx;

    // Seed shared note
    const baseText = [
      "# Sprint Plan",
      "Owner: Unassigned",
      "Budget: 10k",
    ].join("\n");
    await aliceVault.create("SprintPlan.md", baseText);
    await aliceSync.syncAll();
    await bobSync.syncAll();

    // 1A. Overlapping conflict on PUSH (HTTP 412 Precondition Failed)
    await aliceVault.modify(
      aliceVault.getAbstractFileByPath("SprintPlan.md"),
      baseText.replace("Owner: Unassigned", "Owner: Alice")
    );
    await aliceSync.pushFile(aliceVault.getAbstractFileByPath("SprintPlan.md"));

    await bobVault.modify(
      bobVault.getAbstractFileByPath("SprintPlan.md"),
      baseText.replace("Owner: Unassigned", "Owner: Bob")
    );
    const pushConflict = await bobSync.pushFile(bobVault.getAbstractFileByPath("SprintPlan.md"));
    assert.equal(pushConflict.pushed, 1);
    assert.equal(pushConflict.merged, 1);
    assert.equal(pushConflict.conflicts, 1);

    const bobMerged = await bobVault.read(bobVault.getAbstractFileByPath("SprintPlan.md"));
    assert.match(
      bobMerged,
      /<<<<<<< Local\nOwner: Bob\n=======\nOwner: Alice\n>>>>>>> Remote \(SharePoint\)/
    );

    // Alice pulls the conflict-marked file cleanly; Bob resolves markers and syncs cleanly
    const alicePull = await aliceSync.syncAll();
    assert.equal(alicePull.pulled, 1);
    assert.equal(await aliceVault.read(aliceVault.getAbstractFileByPath("SprintPlan.md")), bobMerged);

    const resolvedText = [
      "# Sprint Plan",
      "Owner: Alice & Bob",
      "Budget: 10k",
    ].join("\n");
    await bobVault.modify(bobVault.getAbstractFileByPath("SprintPlan.md"), resolvedText);
    await bobSync.syncAll();
    await aliceSync.syncAll();
    assert.equal(await aliceVault.read(aliceVault.getAbstractFileByPath("SprintPlan.md")), resolvedText);

    // 1B. Pull-time (syncAll) non-overlapping 3-way merge AND overlapping 3-way conflict
    await aliceVault.modify(
      aliceVault.getAbstractFileByPath("SprintPlan.md"),
      resolvedText.replace("# Sprint Plan", "# Sprint Plan Q4")
    );
    await aliceSync.syncAll();

    // Bob modifies a different line locally and runs syncAll() (pull runs before push!)
    await bobVault.modify(
      bobVault.getAbstractFileByPath("SprintPlan.md"),
      resolvedText.replace("Budget: 10k", "Budget: 25k")
    );
    const pullCleanMerge = await bobSync.syncAll();
    assert.equal(pullCleanMerge.merged, 1);
    assert.equal(pullCleanMerge.conflicts, 0);
    assert.equal(pullCleanMerge.pushed, 0, "Already uploaded merged result during pullRemoteChanges");
    const expectedCombined = "# Sprint Plan Q4\nOwner: Alice & Bob\nBudget: 25k";
    assert.equal(await bobVault.read(bobVault.getAbstractFileByPath("SprintPlan.md")), expectedCombined);
    await aliceSync.syncAll();
    assert.equal(await aliceVault.read(aliceVault.getAbstractFileByPath("SprintPlan.md")), expectedCombined);

    // Now trigger overlapping conflict during pullRemoteChanges() (syncAll)
    await aliceVault.modify(
      aliceVault.getAbstractFileByPath("SprintPlan.md"),
      expectedCombined.replace("Budget: 25k", "Budget: 30k (Alice)")
    );
    await aliceSync.syncAll();

    await bobVault.modify(
      bobVault.getAbstractFileByPath("SprintPlan.md"),
      expectedCombined.replace("Budget: 25k", "Budget: 40k (Bob)")
    );
    const pullOverlap = await bobSync.syncAll();
    assert.equal(pullOverlap.merged, 1);
    assert.equal(pullOverlap.conflicts, 1);
    const pullOverlapText = await bobVault.read(bobVault.getAbstractFileByPath("SprintPlan.md"));
    assert.match(
      pullOverlapText,
      /<<<<<<< Local\nBudget: 40k \(Bob\)\n=======\nBudget: 30k \(Alice\)\n>>>>>>> Remote \(SharePoint\)/
    );

    // Resolve 1B conflict on both vaults first so both start 1C from the same base
    const settledBefore1C = "# Sprint Plan Q4\nOwner: Alice & Bob\nBudget: 32k";
    await bobVault.modify(bobVault.getAbstractFileByPath("SprintPlan.md"), settledBefore1C);
    await bobSync.syncAll();
    await aliceSync.syncAll();

    // 1C. Identical concurrent edit on both sides produces 0 conflicts and updates stored eTag cleanly
    const cleanState = "# Sprint Plan Q4\nOwner: Alice & Bob\nBudget: 35k (Agreed)";
    await aliceVault.modify(aliceVault.getAbstractFileByPath("SprintPlan.md"), cleanState);
    await aliceSync.syncAll();
    await bobVault.modify(bobVault.getAbstractFileByPath("SprintPlan.md"), cleanState);
    const identicalRes = await bobSync.syncAll();
    assert.equal(identicalRes.conflicts, 0);
    assert.equal(identicalRes.merged, 0);
    assert.equal(identicalRes.pushed, 0, "Identical remote edit fast-forwards eTag without redundant push");
    assert.equal(ctx.bobStore.files["SprintPlan.md"].eTag, ctx.aliceStore.files["SprintPlan.md"].eTag);
  } finally {
    await ctx.mockGraph.stop();
  }
});

test("Conflict 2: Add/Add text conflicts (both offline syncAll and live pushFile) preserve both authors' notes without overwriting", async () => {
  const ctx = await setupTwoCollaborators();
  try {
    const { aliceVault, bobVault, aliceSync, bobSync } = ctx;

    // 2A. Offline Add/Add via syncAll(): both create Meeting.md independently before syncing
    await aliceVault.create("Notes/Meeting.md", "# Kickoff\n- Alice action item");
    await bobVault.create("Notes/Meeting.md", "# Kickoff\n- Bob action item");

    await aliceSync.syncAll();
    const bobAddAdd = await bobSync.syncAll();
    assert.equal(bobAddAdd.merged, 1);
    assert.equal(bobAddAdd.conflicts, 1);

    const mergedMeeting = await bobVault.read(bobVault.getAbstractFileByPath("Notes/Meeting.md"));
    assert.ok(mergedMeeting.includes("- Bob action item"), "Bob's newly created note must not be lost");
    assert.ok(mergedMeeting.includes("- Alice action item"), "Alice's newly created note must not be lost");
    assert.ok(mergedMeeting.includes("<<<<<<< Local"));

    await aliceSync.syncAll();
    assert.equal(await aliceVault.read(aliceVault.getAbstractFileByPath("Notes/Meeting.md")), mergedMeeting);

    // 2B. Live pushFile() Add/Add race BEFORE /delta has run on Bob's client
    await aliceVault.create("QuickIdea.md", "Alice's remote idea");
    await aliceSync.pushFile(aliceVault.getAbstractFileByPath("QuickIdea.md"));

    // Bob has NOT run pullRemoteChanges() yet, creates QuickIdea.md locally, and triggers live pushFile()
    const bobQuickFile = await bobVault.create("QuickIdea.md", "Bob's local idea");
    const liveRace = await bobSync.pushFile(bobQuickFile);
    assert.equal(liveRace.pushed, 1);
    assert.equal(liveRace.merged, 1);
    assert.equal(liveRace.conflicts, 1);

    const quickMerged = await bobVault.read(bobQuickFile);
    assert.ok(quickMerged.includes("Bob's local idea"));
    assert.ok(quickMerged.includes("Alice's remote idea"));

    // 2C. Identical Add/Add adopts remote ETag with 0 conflicts
    await aliceVault.create("SharedTemplate.md", "# Daily Note\n");
    await aliceSync.syncAll();
    await bobVault.create("SharedTemplate.md", "# Daily Note\n");
    const identicalAdd = await bobSync.syncAll();
    assert.equal(identicalAdd.conflicts, 0);
    assert.equal(identicalAdd.merged, 0);
  } finally {
    await ctx.mockGraph.stop();
  }
});

test("Conflict 3: Binary attachment conflicts (Pull, Push 412, and Add/Add) create conflict copies without losing bytes", async () => {
  const ctx = await setupTwoCollaborators();
  try {
    const { aliceVault, bobVault, aliceSync, bobSync } = ctx;

    const v1Bytes = [0x89, 0x50, 0x4e, 0x47, 0x01, 0x01];
    const aliceV2Bytes = [0x89, 0x50, 0x4e, 0x47, 0x02, 0xaa];
    const bobV2Bytes = [0x89, 0x50, 0x4e, 0x47, 0x02, 0xbb];

    await aliceVault.createBinary("diagrams/arch.png", toBytes(v1Bytes));
    await aliceSync.syncAll();
    await bobSync.syncAll();

    // 3A. Pull-time binary conflict (both modified arch.png offline -> Bob runs syncAll)
    await aliceVault.modifyBinary(aliceVault.getAbstractFileByPath("diagrams/arch.png"), toBytes(aliceV2Bytes));
    await aliceSync.syncAll();

    await bobVault.modifyBinary(bobVault.getAbstractFileByPath("diagrams/arch.png"), toBytes(bobV2Bytes));
    const bobPullBin = await bobSync.syncAll();
    assert.equal(bobPullBin.conflicts, 1);
    assert.equal(bobPullBin.pushed, 1, "Conflict copy should be pushed to SharePoint");

    const bobFiles = bobVault.getFiles().map((f) => f.path);
    const conflictCopyPath = bobFiles.find((p) => p.startsWith("diagrams/arch (Conflict ") && p.endsWith(".png"));
    assert.ok(conflictCopyPath, "Expected diagrams/arch (Conflict ...).png to be created");

    const mainBytesOnBob = Array.from(
      new Uint8Array(await bobVault.readBinary(bobVault.getAbstractFileByPath("diagrams/arch.png")))
    );
    const conflictBytesOnBob = Array.from(
      new Uint8Array(await bobVault.readBinary(bobVault.getAbstractFileByPath(conflictCopyPath)))
    );
    assert.deepEqual(mainBytesOnBob, aliceV2Bytes, "Main file tracks latest remote binary");
    assert.deepEqual(conflictBytesOnBob, bobV2Bytes, "Conflict copy preserves Bob's binary bytes");

    // Alice also receives Bob's conflict copy on her next sync!
    await aliceSync.syncAll();
    assert.ok(aliceVault.getAbstractFileByPath(conflictCopyPath));

    // 3B. Push-time 412 binary conflict (Bob pushes stale ETag directly before /delta runs)
    const aliceV3Bytes = [0x89, 0x50, 0x4e, 0x47, 0x03, 0xcc];
    const bobV3Bytes = [0x89, 0x50, 0x4e, 0x47, 0x03, 0xdd];
    await aliceVault.modifyBinary(aliceVault.getAbstractFileByPath("diagrams/arch.png"), toBytes(aliceV3Bytes));
    await aliceSync.pushFile(aliceVault.getAbstractFileByPath("diagrams/arch.png"));

    await bobVault.modifyBinary(bobVault.getAbstractFileByPath("diagrams/arch.png"), toBytes(bobV3Bytes));
    const bobPushBin = await bobSync.pushFile(bobVault.getAbstractFileByPath("diagrams/arch.png"));
    assert.equal(bobPushBin.conflicts, 1);

    const pushConflictCopies = bobVault
      .getFiles()
      .map((f) => f.path)
      .filter((p) => p.startsWith("diagrams/arch (Conflict ") && p !== conflictCopyPath);
    assert.equal(pushConflictCopies.length, 1);
    assert.deepEqual(
      Array.from(new Uint8Array(await bobVault.readBinary(bobVault.getAbstractFileByPath(pushConflictCopies[0])))),
      bobV3Bytes
    );

    // 3C. Add/Add binary conflict (two offline authors create logo.png with different bytes)
    await aliceVault.createBinary("assets/logo.png", toBytes([1, 2, 3, 4]));
    await bobVault.createBinary("assets/logo.png", toBytes([9, 8, 7, 6]));
    await aliceSync.syncAll();
    const addAddBin = await bobSync.syncAll();
    assert.equal(addAddBin.conflicts, 1);
    const logoConflict = bobVault
      .getFiles()
      .map((f) => f.path)
      .find((p) => p.startsWith("assets/logo (Conflict ") && p.endsWith(".png"));
    assert.ok(logoConflict, "Add/Add binary conflict must preserve local bytes in a conflict copy");
  } finally {
    await ctx.mockGraph.stop();
  }
});

test("Conflict 4: Modify/Delete & Delete/Modify conflicts (files and folders, offline syncAll and live watchers) never destroy unsynced edits", async () => {
  const ctx = await setupTwoCollaborators();
  try {
    const { aliceVault, bobVault, aliceSync, bobSync } = ctx;

    // Seed shared files
    await aliceVault.create("Proposal.md", "v1 baseline");
    await aliceVault.create("Specs/KeepEdited.md", "spec v1");
    await aliceVault.create("Specs/DeleteClean.md", "unchanged v1");
    await aliceVault.create("Hotfix.md", "hotfix v1");
    await aliceVault.create("Architecture.md", "arch v1");
    await aliceSync.syncAll();
    await bobSync.syncAll();

    // 4A. Remote deletes Proposal.md while Bob has unsynced local edits -> Bob's edits preserved & re-uploaded!
    await aliceVault.delete(aliceVault.getAbstractFileByPath("Proposal.md"));
    await aliceSync.deleteLocalPath("Proposal.md", false);

    await bobVault.modify(bobVault.getAbstractFileByPath("Proposal.md"), "v1 baseline\nBob's critical section");
    const modDelRes = await bobSync.syncAll();
    assert.equal(modDelRes.conflicts, 1);
    assert.equal(modDelRes.pushed, 1, "Preserved local file should be re-uploaded to SharePoint");
    assert.equal(
      await bobVault.read(bobVault.getAbstractFileByPath("Proposal.md")),
      "v1 baseline\nBob's critical section"
    );

    await aliceSync.syncAll();
    assert.equal(
      await aliceVault.read(aliceVault.getAbstractFileByPath("Proposal.md")),
      "v1 baseline\nBob's critical section",
      "Alice receives Bob's resurrected Proposal.md"
    );

    // 4B. Remote deletes entire folder Specs/ while Bob edited Specs/KeepEdited.md locally
    await aliceVault.delete(aliceVault.getAbstractFileByPath("Specs"));
    await aliceSync.deleteLocalPath("Specs", true);

    await bobVault.modify(bobVault.getAbstractFileByPath("Specs/KeepEdited.md"), "spec v2 edited by Bob");
    const folderModDel = await bobSync.syncAll();
    assert.equal(folderModDel.conflicts, 1);
    assert.equal(bobVault.getAbstractFileByPath("Specs/DeleteClean.md"), null, "Unmodified file in deleted folder removed");
    assert.ok(bobVault.getAbstractFileByPath("Specs/KeepEdited.md"), "Modified file in deleted folder preserved");
    assert.equal(
      await bobVault.read(bobVault.getAbstractFileByPath("Specs/KeepEdited.md")),
      "spec v2 edited by Bob"
    );

    // 4C. Live Push 412 Modify/Delete: Alice deletes Hotfix.md remotely; before /delta runs, Bob edits & calls pushFile()
    await aliceVault.delete(aliceVault.getAbstractFileByPath("Hotfix.md"));
    await aliceSync.deleteLocalPath("Hotfix.md", false);

    await bobVault.modify(bobVault.getAbstractFileByPath("Hotfix.md"), "hotfix v2 urgent");
    const livePushModDel = await bobSync.pushFile(bobVault.getAbstractFileByPath("Hotfix.md"));
    assert.equal(livePushModDel.pushed, 1);
    assert.equal(livePushModDel.conflicts, 1);
    await aliceSync.syncAll();
    assert.equal(await aliceVault.read(aliceVault.getAbstractFileByPath("Hotfix.md")), "hotfix v2 urgent");

    // 4D. Live Delete/Modify: Alice updates Architecture.md remotely (v2); before /delta runs, Bob deletes Architecture.md locally!
    await aliceVault.modify(aliceVault.getAbstractFileByPath("Architecture.md"), "arch v2 updated by Alice");
    await aliceSync.pushFile(aliceVault.getAbstractFileByPath("Architecture.md"));

    await bobVault.delete(bobVault.getAbstractFileByPath("Architecture.md"));
    const deletedCount = await bobSync.deleteLocalPath("Architecture.md", false);
    assert.equal(deletedCount, 0, "Stale delete must be blocked by If-Match 412 when colleague updated file");
    assert.equal(
      await bobVault.read(bobVault.getAbstractFileByPath("Architecture.md")),
      "arch v2 updated by Alice",
      "Colleague's updated remote note should be restored locally"
    );

    // 4E. Offline Delete/Modify: Alice updates Architecture.md (v3); Bob deletes Architecture.md offline, then runs syncAll()
    await aliceVault.modify(aliceVault.getAbstractFileByPath("Architecture.md"), "arch v3 offline update by Alice");
    await aliceSync.syncAll();

    await bobVault.delete(bobVault.getAbstractFileByPath("Architecture.md"));
    const offlineDelMod = await bobSync.syncAll();
    assert.equal(offlineDelMod.conflicts, 1);
    assert.equal(
      await bobVault.read(bobVault.getAbstractFileByPath("Architecture.md")),
      "arch v3 offline update by Alice"
    );
  } finally {
    await ctx.mockGraph.stop();
  }
});

test("Conflict 5: File rename, folder rename/move, and move to ignored .trash propagate cleanly across collaborators", async () => {
  const ctx = await setupTwoCollaborators();
  try {
    const { aliceVault, bobVault, aliceSync, bobSync } = ctx;

    await aliceVault.create("Drafts/Idea.md", "# Great Idea");
    await aliceVault.create("Drafts/Sub/Notes.md", "Detailed notes");
    await aliceSync.syncAll();
    await bobSync.syncAll();

    // 5A. Single file rename/move: Drafts/Idea.md -> Published/Idea-Final.md
    const movedFile = await bobVault.renameFile("Drafts/Idea.md", "Published/Idea-Final.md");
    await bobSync.renameLocalPath("Drafts/Idea.md", movedFile);

    await aliceSync.syncAll();
    assert.equal(aliceVault.getAbstractFileByPath("Drafts/Idea.md"), null);
    assert.equal(
      await aliceVault.read(aliceVault.getAbstractFileByPath("Published/Idea-Final.md")),
      "# Great Idea"
    );

    // 5B. Whole folder rename/move: Drafts -> Archive/2026
    const movedFolder = await bobVault.renameFolder("Drafts", "Archive/2026");
    await bobSync.renameLocalPath("Drafts", movedFolder);

    await aliceSync.syncAll();
    assert.equal(aliceVault.getAbstractFileByPath("Drafts/Sub/Notes.md"), null);
    assert.equal(
      await aliceVault.read(aliceVault.getAbstractFileByPath("Archive/2026/Sub/Notes.md")),
      "Detailed notes"
    );

    // 5C. Moving a synced file into Obsidian's ignored .trash/ folder deletes it remotely without uploading .trash/
    const trashedFile = await bobVault.renameFile("Published/Idea-Final.md", ".trash/Idea-Final.md");
    await bobSync.renameLocalPath("Published/Idea-Final.md", trashedFile);

    await aliceSync.syncAll();
    assert.equal(aliceVault.getAbstractFileByPath("Published/Idea-Final.md"), null);
    assert.equal(aliceVault.getAbstractFileByPath(".trash/Idea-Final.md"), null);
  } finally {
    await ctx.mockGraph.stop();
  }
});

test("Conflict 6: Three-collaborator double-race (412 during pullRemoteChanges upload) preserves all three authors' edits", async () => {
  const ctx = await setupTwoCollaborators();
  try {
    const { GraphVaultClient, CollaborativeSyncEngine } = await loadSyncModules();
    const { aliceVault, bobVault, aliceSync, bobSync, target } = ctx;
    const port = ctx.mockGraph.server.address().port;
    const baseUrl = `http://127.0.0.1:${port}/v1.0`;

    const carolToken = makeJwt("00000000-0000-4000-8000-000000000103", "Carol De Vries", "carol@contoso.example.com");
    await ctx.aliceGraph.inviteCollaborators(target, ["carol@contoso.example.com"], "write");
    const carolGraph = new GraphVaultClient(async () => carolToken, () => baseUrl);
    const carolVault = new MockVault();
    const carolStore = { files: {}, folders: {} };
    const carolSync = new CollaborativeSyncEngine(carolVault, carolGraph, target, carolStore, async () => {});

    const initial = [
      "# Release Checklist",
      "- Auth: Pending",
      "",
      "- Storage: Pending",
      "",
      "- UI: Pending",
    ].join("\n");
    await aliceVault.create("Checklist.md", initial);
    await aliceSync.syncAll();
    await bobSync.syncAll();
    await carolSync.syncAll();

    // 1. Alice updates Auth and pushes (v2)
    await aliceVault.modify(
      aliceVault.getAbstractFileByPath("Checklist.md"),
      initial.replace("- Auth: Pending", "- Auth: Done (Alice)")
    );
    await aliceSync.pushFile(aliceVault.getAbstractFileByPath("Checklist.md"));

    // 2. Bob edits Storage locally; Carol edits UI locally
    await bobVault.modify(
      bobVault.getAbstractFileByPath("Checklist.md"),
      initial.replace("- Storage: Pending", "- Storage: Done (Bob)")
    );
    await carolVault.modify(
      carolVault.getAbstractFileByPath("Checklist.md"),
      initial.replace("- UI: Pending", "- UI: Done (Carol)")
    );

    // Intercept Bob's first merged PUT during pullRemoteChanges() so Carol slips in v3 right before Bob uploads!
    const origUpload = ctx.bobGraph.uploadFileText.bind(ctx.bobGraph);
    let injectedCarolRace = false;
    ctx.bobGraph.uploadFileText = async (...args) => {
      if (!injectedCarolRace) {
        injectedCarolRace = true;
        // Carol pushes (she hits 412 against Alice's v2, auto-merges Alice + Carol -> v3!)
        await carolSync.pushFile(carolVault.getAbstractFileByPath("Checklist.md"));
      }
      return origUpload(...args);
    };

    const bobRes = await bobSync.syncAll();
    assert.equal(bobRes.merged, 1);
    assert.equal(bobRes.conflicts, 0);

    await aliceSync.syncAll();
    await carolSync.syncAll();

    const finalDoc = await aliceVault.read(aliceVault.getAbstractFileByPath("Checklist.md"));
    assert.ok(finalDoc.includes("- Auth: Done (Alice)"), "Preserved Alice's edit");
    assert.ok(finalDoc.includes("- Storage: Done (Bob)"), "Preserved Bob's edit");
    assert.ok(finalDoc.includes("- UI: Done (Carol)"), "Preserved Carol's edit");
    assert.equal(await bobVault.read(bobVault.getAbstractFileByPath("Checklist.md")), finalDoc);
    assert.equal(await carolVault.read(carolVault.getAbstractFileByPath("Checklist.md")), finalDoc);
  } finally {
    await ctx.mockGraph.stop();
  }
});
