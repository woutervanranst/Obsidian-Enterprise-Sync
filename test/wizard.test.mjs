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
      list: async (dir) => {
        const clean = dir.replace(/^\/+|\/+$/g, "");
        const prefix = clean ? `${clean}/` : "";
        const files = [];
        const folders = [];
        for (const k of this.files.keys()) {
          if (k.startsWith(prefix)) {
            const rest = k.slice(prefix.length);
            if (rest && !rest.includes("/")) files.push(k);
          }
        }
        for (const k of this.folders.keys()) {
          if (k.startsWith(prefix)) {
            const rest = k.slice(prefix.length);
            if (rest && !rest.includes("/")) folders.push(k);
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
  on() {
    return {};
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

const notices = [];

function makeMockElement(tag = "div", opts = {}) {
  const el = {
    tag,
    cls: opts.cls || "",
    text: opts.text || "",
    children: [],
    listeners: {},
    empty() {
      this.children = [];
    },
    addClass(c) {
      this.cls = `${this.cls} ${c}`.trim();
    },
    createDiv(childOpts = {}) {
      const child = makeMockElement("div", childOpts);
      this.children.push(child);
      return child;
    },
    createEl(childTag, childOpts = {}) {
      const child = makeMockElement(childTag, childOpts);
      this.children.push(child);
      return child;
    },
    setText(t) {
      this.text = t;
    },
    addEventListener(type, fn) {
      this.listeners[type] = this.listeners[type] || [];
      this.listeners[type].push(fn);
    },
    async click() {
      for (const fn of this.listeners.click || []) {
        await fn();
      }
    },
  };
  return el;
}

const obsidianMock = {
  TFile: MockTFile,
  TFolder: MockTFolder,
  Vault: MockVault,
  Plugin: class {
    constructor(app, manifest) {
      this.app = app;
      this.manifest = manifest;
    }
    addSettingTab(tab) {
      this._settingTab = tab;
    }
  },
  PluginSettingTab: class {
    constructor(app, plugin) {
      this.app = app;
      this.plugin = plugin;
      this.containerEl = makeMockElement("div");
    }
  },
  Modal: class {
    constructor(app) {
      this.app = app;
      this.contentEl = makeMockElement("div");
    }
    open() {
      if (typeof this.onOpen === "function") this.onOpen();
    }
    close() {
      if (typeof this.onClose === "function") this.onClose();
    }
  },
  Setting: class {
    constructor(containerEl) {
      this.containerEl = containerEl;
      this.buttons = [];
      this.texts = [];
      this.toggles = [];
      this.dropdowns = [];
      if (containerEl && Array.isArray(containerEl.children)) {
        containerEl.children.push(this);
      }
    }
    setName(name) {
      this.name = name;
      return this;
    }
    setDesc(desc) {
      this.desc = desc;
      return this;
    }
    setHeading() {
      this.isHeading = true;
      return this;
    }
    addText(cb) {
      const textComp = {
        inputEl: { type: "text" },
        value: "",
        placeholder: "",
        setValue(v) {
          this.value = v;
          return this;
        },
        setPlaceholder(p) {
          this.placeholder = p;
          return this;
        },
        onChange(fn) {
          this.onChangeFn = fn;
          return this;
        },
        async type(val) {
          this.value = val;
          if (this.onChangeFn) await this.onChangeFn(val);
        },
      };
      cb(textComp);
      this.texts.push(textComp);
      return this;
    }
    addButton(cb) {
      const btnComp = {
        buttonText: "",
        cta: false,
        setButtonText(t) {
          this.buttonText = t;
          return this;
        },
        setCta() {
          this.cta = true;
          return this;
        },
        onClick(fn) {
          this.onClickFn = fn;
          return this;
        },
        async click() {
          if (this.onClickFn) await this.onClickFn();
        },
      };
      cb(btnComp);
      this.buttons.push(btnComp);
      return this;
    }
    addToggle(cb) {
      const toggleComp = {
        value: false,
        setValue(v) {
          this.value = v;
          return this;
        },
        onChange(fn) {
          this.onChangeFn = fn;
          return this;
        },
        async toggle(val) {
          this.value = val;
          if (this.onChangeFn) await this.onChangeFn(val);
        },
      };
      cb(toggleComp);
      this.toggles.push(toggleComp);
      return this;
    }
    addDropdown(cb) {
      const dropComp = {
        options: {},
        addOption(k, v) {
          this.options[k] = v;
          return this;
        },
        setValue(v) {
          this.value = v;
          return this;
        },
        onChange(fn) {
          this.onChangeFn = fn;
          return this;
        },
      };
      cb(dropComp);
      this.dropdowns.push(dropComp);
      return this;
    }
  },
  Notice: class {
    constructor(message) {
      this.message = message;
      notices.push(message);
    }
  },
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

// Compile plugin bundle once for all wizard combination tests
const buildResult = await esbuild.build({
  entryPoints: ["src/main.ts", "src/graph.ts"],
  bundle: true,
  format: "cjs",
  platform: "node",
  external: ["obsidian"],
  write: false,
  outdir: "out",
});

function loadModule(suffix) {
  const file = buildResult.outputFiles.find((f) => f.path.endsWith(suffix));
  const m = { exports: {} };
  const fn = new Function("exports", "require", "module", file.text);
  fn(m.exports, (id) => (id === "obsidian" ? obsidianMock : origLoad(id)), m);
  return m.exports;
}

const { default: EntraSharePointSyncPlugin, DEFAULT_CONSUMER_CLIENT_ID } = loadModule("main.js");
const { GraphVaultClient } = loadModule("graph.js");

// Recursive UI tree inspection helpers
function collectSettings(node) {
  const out = [];
  for (const child of node.children || []) {
    if (child instanceof obsidianMock.Setting) {
      out.push(child);
    } else if (child && Array.isArray(child.children)) {
      out.push(...collectSettings(child));
    }
  }
  return out;
}

function collectElements(node, predicate) {
  const out = [];
  if (predicate(node)) out.push(node);
  for (const child of node.children || []) {
    if (!(child instanceof obsidianMock.Setting) && child) {
      out.push(...collectElements(child, predicate));
    }
  }
  return out;
}

function findSettingByName(tab, name) {
  return collectSettings(tab.containerEl).find((s) => s.name === name);
}

function findButton(tab, buttonText) {
  for (const s of collectSettings(tab.containerEl)) {
    for (const b of s.buttons) {
      if (b.buttonText === buttonText) return b;
    }
  }
  return null;
}

function findChoiceCard(tab, titleText) {
  return collectElements(
    tab.containerEl,
    (el) =>
      el.cls?.includes("entra-sync-choice-card") &&
      el.children.some((c) => c.text === titleText)
  )[0];
}

function findActionTab(tab, titleText) {
  return collectElements(
    tab.containerEl,
    (el) =>
      el.cls?.includes("entra-sync-action-tab") &&
      !el.cls?.includes("entra-sync-action-tabs") &&
      el.children.some((c) => c.text === titleText)
  )[0];
}

function findTargetPill(tab, pillText) {
  return collectElements(
    tab.containerEl,
    (el) => el.cls?.includes("entra-sync-target-pill") && el.text === pillText
  )[0];
}

function findVaultRowButton(tab, cleanTitle) {
  const rows = collectElements(tab.containerEl, (el) =>
    el.cls?.includes("entra-sync-vault-row")
  );
  for (const row of rows) {
    const hasTitle = collectElements(row, (el) => el.text === cleanTitle).length > 0;
    if (hasTitle) {
      const settings = collectSettings(row);
      if (settings[0]?.buttons[0]) return settings[0].buttons[0];
    }
  }
  return null;
}

async function flushMicrotasks() {
  for (let i = 0; i < 6; i++) {
    await new Promise((r) => setTimeout(r, 5));
  }
}

async function bootPlugin({ initialData = {}, port } = {}) {
  let savedData = structuredClone(initialData);
  const openedUrls = [];
  let clipboardText = "";
  const protocolHandlers = {};

  globalThis.window = {
    open(url) {
      openedUrls.push(url);
    },
    setInterval: globalThis.setInterval.bind(globalThis),
    clearInterval: globalThis.clearInterval.bind(globalThis),
    setTimeout: globalThis.setTimeout.bind(globalThis),
    clearTimeout: globalThis.clearTimeout.bind(globalThis),
  };
  Object.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: {
      clipboard: {
        async writeText(t) {
          clipboardText = t;
        },
      },
    },
  });

  const vault = new MockVault();
  const app = { vault, workspace: { getActiveFile: () => null } };
  const plugin = new EntraSharePointSyncPlugin(app, {
    id: "entra-sharepoint-sync",
    name: "OneDrive + Enterprise SharePoint + Teams Sync",
    version: "1.0.8",
  });

  const baseData = port
    ? {
        graphBaseUrl: `http://127.0.0.1:${port}/v1.0`,
        authorityBaseUrl: `http://127.0.0.1:${port}`,
        autoSyncIntervalSec: 0,
        ...savedData,
      }
    : { autoSyncIntervalSec: 0, ...savedData };
  savedData = structuredClone(baseData);

  plugin.loadData = async () => savedData;
  plugin.saveData = async (d) => {
    savedData = structuredClone(d);
  };
  plugin.addStatusBarItem = () => ({
    addClass() {},
    addEventListener() {},
    setText(t) {
      this.text = t;
    },
  });
  plugin.registerObsidianProtocolHandler = (action, handler) => {
    protocolHandlers[action] = handler;
  };
  plugin.addRibbonIcon = () => {};
  plugin.registerEvent = () => {};
  plugin.addCommand = () => {};
  plugin.registerInterval = (id) => id;

  await plugin.onload();
  return {
    plugin,
    vault,
    tab: plugin.settingTab,
    openedUrls,
    getClipboard: () => clipboardText,
    getSavedData: () => savedData,
    triggerProtocol: async (params) => {
      await protocolHandlers["entra-sync-auth"](params);
    },
  };
}

async function signInViaWizardConsumer(ctx, mockGraph, user) {
  ctx.tab.display();
  const consumerCard = findChoiceCard(ctx.tab, "Consumer (OneDrive)");
  assert.ok(consumerCard, "Step 1 must render Consumer (OneDrive) card");
  await consumerCard.click();

  const continueBtn = findButton(ctx.tab, "Continue");
  assert.ok(continueBtn, "Step 1 must render Continue button");
  await continueBtn.click();

  // Verify Consumer Step 2 hides Enterprise-only and dropped options
  assert.equal(findSettingByName(ctx.tab, "Directory (tenant) ID"), undefined);
  assert.equal(findSettingByName(ctx.tab, "Application (client) ID"), undefined);
  assert.equal(findSettingByName(ctx.tab, "Enable SharePoint and Microsoft Teams sites"), undefined);
  assert.equal(findSettingByName(ctx.tab, "Microsoft Graph base URL"), undefined);
  assert.equal(findButton(ctx.tab, "Sign in (Entra credentials)"), null);

  const signInBtn = findButton(ctx.tab, "Sign in with browser");
  assert.ok(signInBtn, "Consumer Step 2 must render 'Sign in with browser'");
  await signInBtn.click();

  assert.equal(ctx.openedUrls.length, 1);
  const authUrl = ctx.openedUrls[0];
  const issued = mockGraph.issueAuthCode(authUrl, user);
  assert.equal(issued.tenantId, "common");
  assert.equal(issued.clientId, DEFAULT_CONSUMER_CLIENT_ID);
  assert.equal(issued.codeChallengeMethod, "S256");
  assert.ok(issued.scope.includes("Files.ReadWrite.All"));
  assert.ok(
    !issued.scope.includes("Sites.ReadWrite.All"),
    "Consumer sign-in must NOT request Sites.ReadWrite.All"
  );

  await ctx.triggerProtocol({ code: issued.code, state: issued.state });
  await flushMicrotasks();
}

async function signInViaWizardEnterprise(ctx, mockGraph, user, tenantId, clientId) {
  ctx.tab.display();
  const entCard = findChoiceCard(ctx.tab, "Enterprise (Teams & SharePoint)");
  assert.ok(entCard, "Step 1 must render Enterprise (Teams & SharePoint) card");
  await entCard.click();
  assert.equal(
    ctx.plugin.settings.includeSitesScope,
    true,
    "Selecting Enterprise automatically sets includeSitesScope = true"
  );

  await findButton(ctx.tab, "Continue").click();

  const tenantSetting = findSettingByName(ctx.tab, "Directory (tenant) ID");
  const clientSetting = findSettingByName(ctx.tab, "Application (client) ID");
  assert.ok(tenantSetting, "Enterprise Step 2 must ask for 'Directory (tenant) ID'");
  assert.ok(clientSetting, "Enterprise Step 2 must ask for 'Application (client) ID'");
  assert.equal(findSettingByName(ctx.tab, "Enable SharePoint and Microsoft Teams sites"), undefined);
  assert.equal(findSettingByName(ctx.tab, "Microsoft Graph base URL"), undefined);
  assert.equal(findButton(ctx.tab, "Sign in (Entra credentials)"), null);

  await tenantSetting.texts[0].type(tenantId);
  await clientSetting.texts[0].type(clientId);

  const beforeCount = ctx.openedUrls.length;
  await findButton(ctx.tab, "Sign in with browser").click();
  assert.equal(ctx.openedUrls.length, beforeCount + 1);

  const authUrl = ctx.openedUrls[ctx.openedUrls.length - 1];
  const issued = mockGraph.issueAuthCode(authUrl, user);
  assert.equal(issued.tenantId, tenantId);
  assert.equal(issued.clientId, clientId);
  assert.ok(
    issued.scope.includes("Sites.ReadWrite.All"),
    "Enterprise sign-in must always include Sites.ReadWrite.All"
  );

  await ctx.triggerProtocol({ code: issued.code, state: issued.state });
  await flushMicrotasks();
}

test("Wizard Combination 1 & 2: Consumer (OneDrive) × PKCE Sign-In × Create Vault (Default & Custom OneDrive Folder)", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const consumerUser = {
    id: "consumer-001",
    displayName: "Clara Consumer",
    email: "clara@outlook.com",
  };

  try {
    // 1A: Default parent folder (ObsidianVaults)
    const ctx1 = await bootPlugin({ port });
    await signInViaWizardConsumer(ctx1, mockGraph, consumerUser);

    await findActionTab(ctx1.tab, "Create vault").click();
    assert.equal(
      findTargetPill(ctx1.tab, "SharePoint or Teams library"),
      undefined,
      "Consumer Create Vault must NOT show SharePoint target pill"
    );

    const parentFolderSetting = findSettingByName(ctx1.tab, "Parent folder in OneDrive");
    assert.ok(parentFolderSetting, "Consumer Create Vault shows 'Parent folder in OneDrive'");
    assert.equal(parentFolderSetting.texts[0].value, "ObsidianVaults");

    await findSettingByName(ctx1.tab, "Vault folder name").texts[0].type("PersonalJournal");
    await ctx1.vault.create("Daily/2026-11-19.md", "# Morning Notes\nSynced via Consumer Wizard.");
    await findButton(ctx1.tab, "Create and connect vault").click();

    assert.ok(ctx1.plugin.settings.activeVault);
    assert.equal(ctx1.plugin.settings.activeVault.name, "PersonalJournal (My OneDrive)");
    // Verify UI transitioned into Connected View
    assert.ok(findButton(ctx1.tab, "Sync now"), "Connected View renders 'Sync now'");
    assert.ok(findButton(ctx1.tab, "Switch vault..."), "Connected View renders 'Switch vault...'");

    // 1B: Custom OneDrive parent folder (Documents/Vaults) from a second device pulling & syncing
    const ctx2 = await bootPlugin({ port });
    await signInViaWizardConsumer(ctx2, mockGraph, consumerUser);
    await findActionTab(ctx2.tab, "Create vault").click();
    await findSettingByName(ctx2.tab, "Parent folder in OneDrive").texts[0].type("Documents/CustomVaults");
    await findSettingByName(ctx2.tab, "Vault folder name").texts[0].type("StudyNotes");
    await ctx2.vault.create("Math.md", "Linear algebra notes");
    await findButton(ctx2.tab, "Create and connect vault").click();

    assert.equal(ctx2.plugin.settings.activeVault.name, "StudyNotes (My OneDrive)");
    assert.equal(ctx2.plugin.settings.oneDriveVaultsFolder, "Documents/CustomVaults");
  } finally {
    await mockGraph.stop();
  }
});

test("Wizard Combination 3: Consumer (OneDrive) × Discover Vaults (My OneDrive + Shared with me + Inline Folder Change)", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const baseUrl = `http://127.0.0.1:${port}/v1.0`;

  const clara = { id: "consumer-clara", displayName: "Clara", email: "clara@outlook.com" };
  const friend = { id: "consumer-dan", displayName: "Dan", email: "dan@live.com" };

  const makeJwt = (u) =>
    `${Buffer.from(JSON.stringify({ alg: "none", typ: "JWT" })).toString("base64url")}.${Buffer.from(
      JSON.stringify({ oid: u.id, name: u.displayName, upn: u.email })
    ).toString("base64url")}.`;

  try {
    // Seed a personal vault in /ObsidianVaults/HomeVault and another in /CustomRoot/SideProject
    const claraGraph = new GraphVaultClient(async () => makeJwt(clara), () => baseUrl);
    await claraGraph.createOrGetVaultFolder("HomeVault", "ObsidianVaults");
    const sideVault = await claraGraph.createOrGetVaultFolder("SideProject", "CustomRoot");
    await claraGraph.uploadFileText(sideVault, "Ideas.md", "# Side Project Ideas");

    // Seed a shared vault from Dan invited to clara@outlook.com
    const danGraph = new GraphVaultClient(async () => makeJwt(friend), () => baseUrl);
    const sharedTrip = await danGraph.createOrGetVaultFolder("FamilyTrip", "ObsidianVaults");
    await danGraph.uploadFileText(sharedTrip, "Itinerary.md", "Day 1: Arrival in Lisbon");
    await danGraph.inviteCollaborators(sharedTrip, [clara.email], "write");

    // Seed a SharePoint site on server to prove Consumer discovery does NOT include SharePoint sites
    mockGraph.seedSharePointSite({
      slug: "CorpOnly",
      displayName: "CorpOnly",
      folders: ["SecretCorpVault"],
    });

    const ctx = await bootPlugin({ port });
    await signInViaWizardConsumer(ctx, mockGraph, clara);

    // Step 3 defaults to Discover vaults
    assert.ok(findVaultRowButton(ctx.tab, "HomeVault"), "Discovers HomeVault in /ObsidianVaults");
    assert.ok(findVaultRowButton(ctx.tab, "FamilyTrip"), "Discovers FamilyTrip from Shared with me");
    assert.equal(
      findVaultRowButton(ctx.tab, "SecretCorpVault"),
      null,
      "Consumer Discover must NOT query followed SharePoint sites"
    );

    // Inline OneDrive folder change -> rescan /CustomRoot
    await findButton(ctx.tab, "Change folder...").click();
    const inlineFolderSetting = findSettingByName(ctx.tab, "OneDrive vaults folder");
    assert.ok(inlineFolderSetting);
    await inlineFolderSetting.texts[0].type("CustomRoot");
    await findButton(ctx.tab, "Save and rescan").click();

    assert.ok(findVaultRowButton(ctx.tab, "SideProject"), "Discovers SideProject after inline folder change");
    assert.ok(findVaultRowButton(ctx.tab, "FamilyTrip"), "Still includes Shared with me vaults");

    // Connect to Dan's shared vault and verify file syncs down
    await findVaultRowButton(ctx.tab, "FamilyTrip").click();
    assert.equal(await ctx.vault.adapter.read("Itinerary.md"), "Day 1: Arrival in Lisbon");
  } finally {
    await mockGraph.stop();
  }
});

test("Wizard Combination 4: Consumer (OneDrive) × Open Vault by URL (1drv.ms Sharing Link & Personal URL)", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const baseUrl = `http://127.0.0.1:${port}/v1.0`;

  const clara = { id: "consumer-url-1", displayName: "Clara", email: "clara@hotmail.com" };
  const dan = { id: "consumer-url-2", displayName: "Dan", email: "dan@outlook.com" };
  const makeJwt = (u) =>
    `${Buffer.from(JSON.stringify({ alg: "none", typ: "JWT" })).toString("base64url")}.${Buffer.from(
      JSON.stringify({ oid: u.id, name: u.displayName, upn: u.email })
    ).toString("base64url")}.`;

  try {
    const danGraph = new GraphVaultClient(async () => makeJwt(dan), () => baseUrl);
    const sharedBookVault = await danGraph.createOrGetVaultFolder("BookClub", "ObsidianVaults");
    await danGraph.uploadFileText(sharedBookVault, "ReadingList.md", "- Dune\n- Neuromancer");
    const shortShareUrl = "https://1drv.ms/f/s!BookClubSharedFolder";
    mockGraph.registerShareUrl(shortShareUrl, sharedBookVault.driveId, sharedBookVault.rootItemId);

    const claraGraph = new GraphVaultClient(async () => makeJwt(clara), () => baseUrl);
    const personalRecipes = await claraGraph.createOrGetVaultFolder("Recipes", "ObsidianVaults");
    await claraGraph.uploadFileText(personalRecipes, "Pasta.md", "Boil salted water.");

    const ctx = await bootPlugin({ port });
    await signInViaWizardConsumer(ctx, mockGraph, clara);

    // 4A: Open 1drv.ms sharing link
    await findActionTab(ctx.tab, "Open vault by URL").click();
    await findSettingByName(ctx.tab, "Vault folder URL").texts[0].type(shortShareUrl);
    await findButton(ctx.tab, "Open vault").click();

    assert.equal(
      ctx.plugin.settings.includeSitesScope,
      false,
      "Opening a Consumer OneDrive URL must NOT flip includeSitesScope to true"
    );
    assert.equal(await ctx.vault.adapter.read("ReadingList.md"), "- Dune\n- Neuromancer");

    // 4B: Switch vault -> Open personal OneDrive folder URL
    await findButton(ctx.tab, "Switch vault...").click();
    await findActionTab(ctx.tab, "Open vault by URL").click();
    await findSettingByName(ctx.tab, "Vault folder URL").texts[0].type(
      "https://my.microsoftpersonalcontent.com/personal/clara_hotmail_com/Documents/ObsidianVaults/Recipes"
    );
    await findButton(ctx.tab, "Open vault").click();
    assert.equal(await ctx.vault.adapter.read("Pasta.md"), "Boil salted water.");
  } finally {
    await mockGraph.stop();
  }
});

test("Wizard Combination 5: Enterprise Sign-In Validation, IT Admin Copy, Back/Forward State Preservation & PKCE Scopes", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const alice = {
    id: "ent-alice-1",
    displayName: "Alice Peeters",
    email: "alice@contoso.onmicrosoft.com",
  };

  try {
    notices.length = 0;
    const ctx = await bootPlugin({ port });
    ctx.tab.display();

    await findChoiceCard(ctx.tab, "Enterprise (Teams & SharePoint)").click();
    await findButton(ctx.tab, "Continue").click();

    // Validation when fields are empty
    await findButton(ctx.tab, "Sign in with browser").click();
    assert.equal(ctx.openedUrls.length, 0, "Must not open browser when Tenant/Client ID are empty");
    assert.ok(
      notices.some((m) =>
        m.includes("Directory (tenant) ID") && m.includes("Application (client) ID")
      )
    );

    // Copy IT admin request template
    await findButton(ctx.tab, "Copy admin request").click();
    assert.ok(ctx.getClipboard().includes("Directory (tenant) ID"));
    assert.ok(ctx.getClipboard().includes("Application (client) ID"));
    assert.ok(ctx.getClipboard().includes("Sites.ReadWrite.All"));
    assert.ok(ctx.getClipboard().includes("obsidian://entra-sync-auth"));

    // Type IDs, go Back to Step 1, toggle Consumer and back to Enterprise, verify IDs preserved!
    await findSettingByName(ctx.tab, "Directory (tenant) ID").texts[0].type(
      "contoso.onmicrosoft.com"
    );
    await findSettingByName(ctx.tab, "Application (client) ID").texts[0].type(
      "11111111-2222-3333-4444-555555555555"
    );

    await findButton(ctx.tab, "Back").click();
    await findChoiceCard(ctx.tab, "Consumer (OneDrive)").click();
    assert.equal(ctx.plugin.settings.includeSitesScope, false);
    await findChoiceCard(ctx.tab, "Enterprise (Teams & SharePoint)").click();
    assert.equal(ctx.plugin.settings.includeSitesScope, true);
    await findButton(ctx.tab, "Continue").click();

    assert.equal(
      findSettingByName(ctx.tab, "Directory (tenant) ID").texts[0].value,
      "contoso.onmicrosoft.com"
    );
    assert.equal(
      findSettingByName(ctx.tab, "Application (client) ID").texts[0].value,
      "11111111-2222-3333-4444-555555555555"
    );

    await findButton(ctx.tab, "Sign in with browser").click();
    assert.equal(ctx.openedUrls.length, 1);
    const issued = mockGraph.issueAuthCode(ctx.openedUrls[0], alice);
    assert.equal(issued.tenantId, "contoso.onmicrosoft.com");
    assert.equal(issued.clientId, "11111111-2222-3333-4444-555555555555");
    assert.ok(issued.scope.includes("Sites.ReadWrite.All"));

    await ctx.triggerProtocol({ code: issued.code, state: issued.state });
    await flushMicrotasks();

    // Auto-advanced to Step 3
    assert.ok(findActionTab(ctx.tab, "Discover vaults"));
    assert.ok(findActionTab(ctx.tab, "Open vault by URL"));
    assert.ok(findActionTab(ctx.tab, "Create vault"));
  } finally {
    await mockGraph.stop();
  }
});

test("Wizard Combination 6: Enterprise × Discover Vaults (Followed SharePoint Sites + Shared with me + OneDrive for Business)", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const baseUrl = `http://127.0.0.1:${port}/v1.0`;

  const alice = {
    id: "ent-alice-disc",
    displayName: "Alice Peeters",
    email: "alice@contoso.com",
  };
  const bob = {
    id: "ent-bob-disc",
    displayName: "Bob Janssens",
    email: "bob@contoso.com",
  };
  const makeJwt = (u) =>
    `${Buffer.from(JSON.stringify({ alg: "none", typ: "JWT" })).toString("base64url")}.${Buffer.from(
      JSON.stringify({ oid: u.id, name: u.displayName, upn: u.email })
    ).toString("base64url")}.`;

  try {
    // 1. Seed followed SharePoint site with folder "EngHandbook" and file inside it
    const seededSite = mockGraph.seedSharePointSite({
      slug: "Engineering",
      displayName: "Engineering Hub",
      folders: ["EngHandbook"],
    });
    const aliceGraph = new GraphVaultClient(async () => makeJwt(alice), () => baseUrl);
    await aliceGraph.uploadFileText(
      {
        name: "EngHandbook",
        driveId: seededSite.driveId,
        rootItemId: seededSite.createdFolders[0].id,
      },
      "ADR-001.md",
      "# ADR-001: Use PKCE everywhere"
    );

    // 2. Seed Alice's OneDrive for Business folder
    await aliceGraph.createOrGetVaultFolder("MyWorkNotes", "ObsidianVaults");

    // 3. Seed Bob's shared folder invited to Alice
    const bobGraph = new GraphVaultClient(async () => makeJwt(bob), () => baseUrl);
    const sharedSec = await bobGraph.createOrGetVaultFolder("SecurityPlaybook", "ObsidianVaults");
    await bobGraph.inviteCollaborators(sharedSec, [alice.email], "write");

    const ctx = await bootPlugin({ port });
    await signInViaWizardEnterprise(
      ctx,
      mockGraph,
      alice,
      "contoso-tenant-uuid",
      "contoso-client-uuid"
    );

    assert.ok(
      findVaultRowButton(ctx.tab, "EngHandbook"),
      "Enterprise Discover lists SharePoint followed site vault"
    );
    assert.ok(
      findVaultRowButton(ctx.tab, "MyWorkNotes"),
      "Enterprise Discover lists OneDrive for Business vault"
    );
    assert.ok(
      findVaultRowButton(ctx.tab, "SecurityPlaybook"),
      "Enterprise Discover lists Shared with me vault"
    );

    await findVaultRowButton(ctx.tab, "EngHandbook").click();
    assert.equal(await ctx.vault.adapter.read("ADR-001.md"), "# ADR-001: Use PKCE everywhere");
  } finally {
    await mockGraph.stop();
  }
});

test("Wizard Combination 7: Enterprise × Open Vault by URL (SharePoint Direct, Forms/AllItems.aspx, Fallback Site API, and Microsoft Teams Channel URL)", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();
  const baseUrl = `http://127.0.0.1:${port}/v1.0`;

  const alice = {
    id: "ent-alice-url",
    displayName: "Alice Peeters",
    email: "alice@contoso.com",
  };
  const makeJwt = (u) =>
    `${Buffer.from(JSON.stringify({ alg: "none", typ: "JWT" })).toString("base64url")}.${Buffer.from(
      JSON.stringify({ oid: u.id, name: u.displayName, upn: u.email })
    ).toString("base64url")}.`;

  try {
    const graph = new GraphVaultClient(async () => makeJwt(alice), () => baseUrl);

    // Seed SharePoint site, fallback SharePoint site, and Teams channel drive
    const engSite = mockGraph.seedSharePointSite({
      slug: "Engineering",
      displayName: "Engineering",
      folders: ["PlatformVault"],
    });
    await graph.uploadFileText(
      {
        name: "PlatformVault",
        driveId: engSite.driveId,
        rootItemId: engSite.createdFolders[0].id,
      },
      "Spec.md",
      "SharePoint Spec v1"
    );

    const secSite = mockGraph.seedSharePointSite({
      slug: "SecOps",
      displayName: "SecOps",
      hostname: "fallback.sharepoint.com",
      folders: ["IncidentVault"],
    });
    mockGraph.setForceShareFailureHost("fallback.sharepoint.com", true);
    await graph.uploadFileText(
      {
        name: "IncidentVault",
        driveId: secSite.driveId,
        rootItemId: secSite.createdFolders[0].id,
      },
      "Runbook.md",
      "Fallback Site Resolution Works"
    );

    const teamsSite = mockGraph.seedSharePointSite({
      slug: "ProductLaunch",
      displayName: "Product Launch Team",
      kind: "teams",
      folders: ["General/LaunchVault"],
    });
    await graph.uploadFileText(
      {
        name: "LaunchVault",
        driveId: teamsSite.driveId,
        rootItemId: teamsSite.createdFolders[0].id,
      },
      "Checklist.md",
      "Teams Channel Vault Synced"
    );

    const ctx = await bootPlugin({ port });
    await signInViaWizardEnterprise(
      ctx,
      mockGraph,
      alice,
      "contoso-tenant-uuid",
      "contoso-client-uuid"
    );

    // 7A: Direct SharePoint URL
    await findActionTab(ctx.tab, "Open vault by URL").click();
    await findSettingByName(ctx.tab, "Vault folder URL").texts[0].type(
      "https://contoso.sharepoint.com/sites/Engineering/Shared%20Documents/PlatformVault"
    );
    await findButton(ctx.tab, "Open vault").click();
    assert.equal(await ctx.vault.adapter.read("Spec.md"), "SharePoint Spec v1");

    // 7B: SharePoint Forms/AllItems.aspx?id=... URL
    await findButton(ctx.tab, "Switch vault...").click();
    await findActionTab(ctx.tab, "Open vault by URL").click();
    await findSettingByName(ctx.tab, "Vault folder URL").texts[0].type(
      "https://contoso.sharepoint.com/sites/Engineering/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FEngineering%2FShared%20Documents%2FPlatformVault&viewid=0000"
    );
    await findButton(ctx.tab, "Open vault").click();
    assert.ok(ctx.plugin.settings.activeVault.name.includes("PlatformVault"));

    // 7C: SharePoint fallback resolution (/sites/{host}:/{path}) when /shares/ returns 403
    await findButton(ctx.tab, "Switch vault...").click();
    await findActionTab(ctx.tab, "Open vault by URL").click();
    await findSettingByName(ctx.tab, "Vault folder URL").texts[0].type(
      "https://fallback.sharepoint.com/sites/SecOps/Shared%20Documents/IncidentVault"
    );
    await findButton(ctx.tab, "Open vault").click();
    assert.equal(await ctx.vault.adapter.read("Runbook.md"), "Fallback Site Resolution Works");

    // 7D: Microsoft Teams channel URL (/teams/ProductLaunch/Shared Documents/General/LaunchVault)
    await findButton(ctx.tab, "Switch vault...").click();
    await findActionTab(ctx.tab, "Open vault by URL").click();
    await findSettingByName(ctx.tab, "Vault folder URL").texts[0].type(
      "https://contoso.sharepoint.com/teams/ProductLaunch/Shared%20Documents/General/LaunchVault"
    );
    await findButton(ctx.tab, "Open vault").click();
    assert.equal(await ctx.vault.adapter.read("Checklist.md"), "Teams Channel Vault Synced");
  } finally {
    await mockGraph.stop();
  }
});

test("Wizard Combination 8 & 9: Enterprise × Create Vault in SharePoint/Teams Library (Root, Nested Subfolder, Existing Folder) AND in My OneDrive for Business", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();

  const alice = {
    id: "ent-alice-create",
    displayName: "Alice Peeters",
    email: "alice@contoso.com",
  };
  const bob = {
    id: "ent-bob-create",
    displayName: "Bob Janssens",
    email: "bob@contoso.com",
  };

  try {
    mockGraph.seedSharePointSite({
      slug: "Engineering",
      displayName: "Engineering",
      folders: [],
    });
    mockGraph.seedSharePointSite({
      slug: "DesignTeam",
      displayName: "Design Team",
      kind: "teams",
      folders: ["General"],
    });

    const ctx = await bootPlugin({ port });
    await signInViaWizardEnterprise(
      ctx,
      mockGraph,
      alice,
      "contoso-tenant-uuid",
      "contoso-client-uuid"
    );

    // 8A: Create vault at SharePoint Document Library root
    await findActionTab(ctx.tab, "Create vault").click();
    const spPill = findTargetPill(ctx.tab, "SharePoint or Teams library");
    const odPill = findTargetPill(ctx.tab, "My OneDrive for Business");
    assert.ok(spPill, "Enterprise Create Vault renders 'SharePoint or Teams library' pill");
    assert.ok(odPill, "Enterprise Create Vault renders 'My OneDrive for Business' pill");

    await spPill.click();
    await findSettingByName(ctx.tab, "Vault folder name").texts[0].type("CloudArchitecture");
    await findSettingByName(
      ctx.tab,
      "Parent SharePoint or Teams library URL"
    ).texts[0].type("https://contoso.sharepoint.com/sites/Engineering/Shared%20Documents");
    await ctx.vault.create("SystemOverview.md", "# Cloud Architecture");
    await findButton(ctx.tab, "Create and connect vault").click();

    assert.equal(
      ctx.plugin.settings.activeVault.name,
      "CloudArchitecture (SharePoint: Engineering)"
    );
    assert.equal(ctx.plugin.settings.activeVault.driveId, "sp-drive-engineering");

    // 8B: Create vault inside a nested Microsoft Teams channel folder (/General -> Q4DesignVault)
    await findButton(ctx.tab, "Switch vault...").click();
    await findActionTab(ctx.tab, "Create vault").click();
    await findTargetPill(ctx.tab, "SharePoint or Teams library").click();
    await findSettingByName(ctx.tab, "Vault folder name").texts[0].type("Q4DesignVault");
    await findSettingByName(
      ctx.tab,
      "Parent SharePoint or Teams library URL"
    ).texts[0].type("https://contoso.sharepoint.com/teams/DesignTeam/Shared%20Documents/General");
    await findButton(ctx.tab, "Create and connect vault").click();

    assert.equal(
      ctx.plugin.settings.activeVault.name,
      "Q4DesignVault (SharePoint: DesignTeam)"
    );
    assert.equal(ctx.plugin.settings.activeVault.driveId, "sp-drive-designteam");
    const createdNestedId = ctx.plugin.settings.activeVault.rootItemId;

    // 8C: Colleague Bob runs Create Vault pointing at the already-existing nested Teams vault (409 fallback)
    const bobCtx = await bootPlugin({ port });
    await signInViaWizardEnterprise(
      bobCtx,
      mockGraph,
      bob,
      "contoso-tenant-uuid",
      "contoso-client-uuid"
    );
    await findActionTab(bobCtx.tab, "Create vault").click();
    await findSettingByName(bobCtx.tab, "Vault folder name").texts[0].type("Q4DesignVault");
    await findSettingByName(
      bobCtx.tab,
      "Parent SharePoint or Teams library URL"
    ).texts[0].type("https://contoso.sharepoint.com/teams/DesignTeam/Shared%20Documents/General");
    await findButton(bobCtx.tab, "Create and connect vault").click();
    assert.equal(
      bobCtx.plugin.settings.activeVault.rootItemId,
      createdNestedId,
      "Existing subfolder inside Teams channel resolves to the exact same item ID"
    );

    // 9: Enterprise Create Vault in My OneDrive for Business
    await findButton(ctx.tab, "Switch vault...").click();
    await findActionTab(ctx.tab, "Create vault").click();
    await findTargetPill(ctx.tab, "My OneDrive for Business").click();

    assert.equal(
      findSettingByName(ctx.tab, "Parent SharePoint or Teams library URL"),
      undefined,
      "Selecting OneDrive for Business hides SharePoint parent URL input"
    );
    const odParent = findSettingByName(ctx.tab, "Parent folder in OneDrive");
    assert.ok(odParent, "Selecting OneDrive for Business shows 'Parent folder in OneDrive'");
    await odParent.texts[0].type("WorkVaults");
    await findSettingByName(ctx.tab, "Vault folder name").texts[0].type("PersonalOKRs");
    await findButton(ctx.tab, "Create and connect vault").click();

    assert.equal(ctx.plugin.settings.activeVault.name, "PersonalOKRs (My OneDrive)");
    assert.equal(ctx.plugin.settings.activeVault.driveId, `drive-${alice.id}`);
  } finally {
    await mockGraph.stop();
  }
});

test("Wizard Combination 10: Connected View Actions (Force Token Refresh, Reset Delta, Cancel Switch Vault, Sign Out) & Legacy v1.0.8 Settings Migration", async () => {
  const mockGraph = createMockGraphServer({ port: 0, host: "127.0.0.1" });
  const port = await mockGraph.start();

  const alice = {
    id: "ent-alice-connected",
    displayName: "Alice Peeters",
    email: "alice@contoso.com",
  };

  try {
    const ctx = await bootPlugin({ port });
    await signInViaWizardConsumer(ctx, mockGraph, alice);
    await findActionTab(ctx.tab, "Create vault").click();
    await findSettingByName(ctx.tab, "Vault folder name").texts[0].type("SteadyStateVault");
    await findButton(ctx.tab, "Create and connect vault").click();

    // Connected view actions: Force token refresh rotates refresh token via PKCE endpoint
    const oldRefreshToken = ctx.plugin.settings.tokens.refreshToken;
    await findButton(ctx.tab, "Force token refresh").click();
    assert.notEqual(
      ctx.plugin.settings.tokens.refreshToken,
      oldRefreshToken,
      "Force token refresh rotates OAuth2 PKCE refresh token"
    );

    // Reset delta and resync
    assert.ok(ctx.plugin.settings.syncStore.deltaLink);
    await findButton(ctx.tab, "Reset delta and resync").click();
    assert.ok(ctx.plugin.settings.syncStore.deltaLink, "Re-establishes deltaLink after reset");

    // Switch vault... -> Cancel returns to Connected View without losing active vault
    await findButton(ctx.tab, "Switch vault...").click();
    const cancelSwitchBtn = findButton(ctx.tab, "Cancel");
    assert.ok(cancelSwitchBtn, "Step 3 shows Cancel button when switching from an active vault");
    await cancelSwitchBtn.click();
    assert.ok(findButton(ctx.tab, "Sync now"), "Cancel returns to Connected View");
    assert.equal(ctx.plugin.settings.activeVault.name, "SteadyStateVault (My OneDrive)");

    // Sign out returns to Wizard Step 1
    await findButton(ctx.tab, "Sign out").click();
    assert.equal(ctx.plugin.settings.tokens, undefined);
    assert.equal(ctx.plugin.settings.activeVault, undefined);
    assert.ok(findChoiceCard(ctx.tab, "Consumer (OneDrive)"), "Sign out returns to Step 1");

    // Legacy v1.0.8 Consumer data.json migration (no `edition` field, default tenant "common")
    const legacyConsumerCtx = await bootPlugin({
      port,
      initialData: {
        tenantId: "common",
        clientId: DEFAULT_CONSUMER_CLIENT_ID,
        includeSitesScope: false,
        oneDriveVaultsFolder: "Documents/LegacyVaults",
      },
    });
    assert.equal(legacyConsumerCtx.plugin.settings.edition, "consumer");
    assert.equal(legacyConsumerCtx.plugin.settings.includeSitesScope, false);
    assert.equal(
      legacyConsumerCtx.plugin.settings.tenantId,
      "",
      "Legacy 'common' tenant is normalized so Enterprise input starts clean"
    );
    assert.equal(legacyConsumerCtx.plugin.settings.oneDriveVaultsFolder, "Documents/LegacyVaults");

    // Legacy v1.0.8 Enterprise data.json migration (no `edition` field, custom tenant/client + custom graphBaseUrl)
    const customGraphUrl = `http://127.0.0.1:${port}/v1.0`;
    const legacyEntCtx = await bootPlugin({
      initialData: {
        tenantId: "72f988bf-86f1-41af-91ab-2d7cd011db47",
        clientId: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        graphBaseUrl: customGraphUrl,
        includeSitesScope: false,
      },
    });
    assert.equal(legacyEntCtx.plugin.settings.edition, "enterprise");
    assert.equal(
      legacyEntCtx.plugin.settings.includeSitesScope,
      true,
      "Migrated Enterprise settings enforce includeSitesScope = true"
    );
    assert.equal(
      legacyEntCtx.plugin.settings.tenantId,
      "72f988bf-86f1-41af-91ab-2d7cd011db47"
    );
    assert.equal(legacyEntCtx.plugin.settings.graphBaseUrl, customGraphUrl);
  } finally {
    await mockGraph.stop();
  }
});
