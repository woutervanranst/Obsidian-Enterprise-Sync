import { App, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile, TFolder } from "obsidian";
import { AuthTokens, EntraPkceAuth } from "./auth";
import { GraphVaultClient, RemoteVaultTarget } from "./graph";
import { CollaborativeSyncEngine, SyncStore, isIgnoredVaultPath } from "./sync";
import { VersionHistoryModal } from "./historyModal";
import { ShareVaultModal } from "./shareModal";
import { SyncLogger, defaultLogger } from "./logger";

export type AccountEdition = "consumer" | "enterprise";

export const DEFAULT_CONSUMER_CLIENT_ID = "ada87c12-a404-484b-a7dc-10c54beb2d7d";
export const DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export interface EnterpriseDirectoryConfiguration {
  emailDomain: string;
  tenantId: string;
  clientId: string;
}

export const ENTERPRISE_DIRECTORY_CONFIGURATIONS: ReadonlyArray<
  EnterpriseDirectoryConfiguration
> = [
  {
    emailDomain: "ae.be",
    tenantId: "65830af0-624d-48a0-a9da-07bd07a31094",
    clientId: "4baf3384-e011-414d-914c-3be68d03f488",
  },
];

export function findEnterpriseDirectoryConfiguration(
  email: string
): EnterpriseDirectoryConfiguration | undefined {
  const normalizedEmail = email.trim().toLowerCase();
  const separator = normalizedEmail.lastIndexOf("@");
  if (separator <= 0 || separator === normalizedEmail.length - 1) {
    return undefined;
  }

  const domain = normalizedEmail.slice(separator + 1);
  return ENTERPRISE_DIRECTORY_CONFIGURATIONS.find(
    (configuration) => configuration.emailDomain.toLowerCase() === domain
  );
}

const ADMIN_REQUEST_TEMPLATE = [
  "Subject: Request for Microsoft Entra ID App Registration — Obsidian SharePoint & OneDrive Sync",
  "",
  "Hi,",
  "",
  "Could you create a Microsoft Entra ID App Registration so we can collaborate on Obsidian vaults stored in SharePoint Online, Microsoft Teams, and OneDrive?",
  "",
  "1. Name: Obsidian SharePoint Sync",
  "2. Supported account types: Accounts in this organizational directory only (Single tenant) (or Multitenant if guests need access)",
  "3. Authentication -> Platform: Mobile and desktop applications -> Custom Redirect URI: obsidian://entra-sync-auth",
  "4. API permissions (Microsoft Graph — Delegated): User.Read, Files.ReadWrite.All, Sites.ReadWrite.All, offline_access (with admin consent)",
  "",
  "Once created, please share back:",
  "- Directory (tenant) ID",
  "- Application (client) ID",
].join("\n");

export interface EntraSyncSettings {
  edition: AccountEdition;
  tenantId: string;
  clientId: string;
  graphBaseUrl: string;
  authorityBaseUrl?: string;
  oneDriveVaultsFolder: string;
  includeSitesScope: boolean;
  syncObsidianConfig: boolean;
  autoSyncIntervalSec: number;
  tokens?: AuthTokens;
  activeVault?: RemoteVaultTarget;
  syncStore: SyncStore;
}

const DEFAULT_SETTINGS: EntraSyncSettings = {
  edition: "consumer",
  tenantId: "",
  clientId: "",
  graphBaseUrl: DEFAULT_GRAPH_BASE_URL,
  oneDriveVaultsFolder: "ObsidianVaults",
  includeSitesScope: false,
  syncObsidianConfig: true,
  autoSyncIntervalSec: 15,
  syncStore: { files: {} },
};

function openExternalUrl(url: string): void {
  if (typeof window !== "undefined" && typeof window.open === "function") {
    window.open(url);
  }
}

export default class EntraSharePointSyncPlugin extends Plugin {
  settings: EntraSyncSettings = { ...DEFAULT_SETTINGS, syncStore: { files: {} } };
  logger: SyncLogger = defaultLogger;
  auth!: EntraPkceAuth;
  graph!: GraphVaultClient;
  syncEngine?: CollaborativeSyncEngine;
  statusBarEl?: HTMLElement;
  settingTab?: EntraSyncSettingTab;
  private syncTimerId: number | null = null;
  private modifyTimeoutId: number | null = null;

  async onload() {
    await this.loadSettings();

    this.statusBarEl = this.addStatusBarItem();
    this.statusBarEl.addClass("entra-sync-status-bar");
    this.statusBarEl.addEventListener("click", () => {
      void this.triggerManualSync();
    });

    this.reinitClients();

    // Handle obsidian://entra-sync-auth?code=...
    this.registerObsidianProtocolHandler("entra-sync-auth", async (params) => {
      if (params.code) {
        try {
          this.settings.tokens = await this.auth.handleCallback(params.code);
          await this.saveSettings();
          this.reinitClients();
          this.settingTab?.onAuthCompleted();
          new Notice(
            `Signed in with Microsoft as ${
              this.settings.tokens.userDisplayName ||
              this.settings.tokens.userEmail ||
              "Microsoft user"
            }!`
          );
        } catch (e) {
          new Notice(`Sign-in failed: ${String(e)}`);
        }
      }
    });

    // Ribbon: Sync now
    this.addRibbonIcon("refresh-cw", "Sync vault now", async () => {
      await this.triggerManualSync();
    });

    // Ribbon: Share vault
    this.addRibbonIcon("users", "Share vault with colleagues", () => {
      if (!this.settings.tokens) {
        new Notice("Please sign in with Microsoft Entra ID first.");
        return;
      }
      if (!this.settings.activeVault) {
        new Notice("Please select or create a shared vault in settings first.");
        return;
      }
      new ShareVaultModal(this.app, this.settings.activeVault, this.graph).open();
    });

    // Ribbon: Version history
    this.addRibbonIcon("history", "View note version history", () => {
      this.openActiveFileHistory();
    });

    // Auto-push when user creates or modifies a note or attachment (debounced)
    const schedulePush = (abstractFile: unknown) => {
      if (
        this.syncEngine &&
        abstractFile instanceof TFile &&
        !isIgnoredVaultPath(
          abstractFile.path,
          this.settings.syncObsidianConfig,
          this.app.vault.configDir
        ) &&
        !this.syncEngine.isPathSuppressed(abstractFile.path)
      ) {
        if (this.modifyTimeoutId !== null) {
          window.clearTimeout(this.modifyTimeoutId);
        }
        this.modifyTimeoutId = window.setTimeout(async () => {
          try {
            await this.syncEngine?.pushFile(abstractFile);
            this.updateStatusBar("Synced");
          } catch (err) {
            this.logger.error(`Auto-push failed for ${abstractFile.path}`, String(err));
            this.updateStatusBar("Sync error");
          }
        }, 800);
      }
    };

    this.registerEvent(this.app.vault.on("create", schedulePush));
    this.registerEvent(this.app.vault.on("modify", schedulePush));

    // Propagate local file & folder deletions to SharePoint / OneDrive
    this.registerEvent(
      this.app.vault.on("delete", (abstractFile) => {
        if (
          this.syncEngine &&
          (abstractFile instanceof TFile || abstractFile instanceof TFolder) &&
          !isIgnoredVaultPath(
            abstractFile.path,
            this.settings.syncObsidianConfig,
            this.app.vault.configDir
          ) &&
          !this.syncEngine.isDeleteSuppressed(abstractFile.path)
        ) {
          const isFolder = abstractFile instanceof TFolder;
          void this.syncEngine
            .deleteLocalPath(abstractFile.path, isFolder)
            .then(() => this.updateStatusBar("Synced"))
            .catch((err) => {
              this.logger.error(
                `Remote delete failed for ${abstractFile.path}`,
                String(err)
              );
              this.updateStatusBar("Sync error");
            });
        }
      })
    );

    // Propagate local file & folder renames/moves to SharePoint / OneDrive
    this.registerEvent(
      this.app.vault.on("rename", (abstractFile, oldPath) => {
        if (
          this.syncEngine &&
          (abstractFile instanceof TFile || abstractFile instanceof TFolder) &&
          !isIgnoredVaultPath(
            abstractFile.path,
            this.settings.syncObsidianConfig,
            this.app.vault.configDir
          )
        ) {
          void this.syncEngine
            .renameLocalPath(oldPath, abstractFile)
            .then(() => this.updateStatusBar("Synced"))
            .catch((err) => {
              this.logger.error(
                `Remote rename failed from ${oldPath} to ${abstractFile.path}`,
                String(err)
              );
              this.updateStatusBar("Sync error");
            });
        }
      })
    );

    // Command: Sync vault now
    this.addCommand({
      id: "entra-sync-now",
      name: "Sync vault now",
      callback: async () => {
        await this.triggerManualSync();
      },
    });

    // Command: Share vault with colleagues
    this.addCommand({
      id: "entra-share-vault",
      name: "Share vault with colleagues",
      callback: () => {
        if (!this.settings.activeVault) {
          new Notice("No active shared vault configured.");
          return;
        }
        new ShareVaultModal(this.app, this.settings.activeVault, this.graph).open();
      },
    });

    // Command: View SharePoint / OneDrive version history for active note
    this.addCommand({
      id: "view-note-version-history",
      name: "View note version history",
      callback: () => {
        this.openActiveFileHistory();
      },
    });

    // Command: View sync diagnostic logs
    this.addCommand({
      id: "entra-view-sync-logs",
      name: "View sync diagnostic logs",
      callback: () => {
        new SyncLogsModal(this.app, this).open();
      },
    });

    this.settingTab = new EntraSyncSettingTab(this.app, this);
    this.addSettingTab(this.settingTab);

    // Perform an initial sync on startup if already signed in and connected to a vault
    if (this.settings.tokens && this.settings.activeVault) {
      window.setTimeout(() => {
        void this.triggerManualSync(true);
      }, 1500);
    }
  }

  onunload() {
    if (this.syncTimerId !== null) {
      window.clearInterval(this.syncTimerId);
      this.syncTimerId = null;
    }
    if (this.modifyTimeoutId !== null) {
      window.clearTimeout(this.modifyTimeoutId);
      this.modifyTimeoutId = null;
    }
  }

  openActiveFileHistory() {
    const file =
      this.app.workspace.getActiveFile() || this.app.vault.getMarkdownFiles()[0];
    if (!file) {
      new Notice("No Markdown note open.");
      return;
    }
    const fileState = this.settings.syncStore.files[file.path];
    if (!fileState || !this.settings.activeVault) {
      new Notice("This note has not been synced to SharePoint / OneDrive yet.");
      return;
    }
    new VersionHistoryModal(
      this.app,
      file,
      this.settings.activeVault.driveId,
      fileState.itemId,
      this.graph,
      async (restoredText) => {
        if (this.syncEngine) {
          await this.syncEngine.pushMarkdownFile(file);
        } else {
          fileState.baseText = restoredText;
          await this.saveSettings();
        }
      }
    ).open();
  }

  async triggerManualSync(silent = false) {
    if (!this.syncEngine || !this.settings.tokens || !this.settings.activeVault) {
      if (!silent) {
        new Notice("Configure Entra ID sign-in and select a vault in Settings first.");
      }
      return;
    }
    try {
      const summary = await this.syncEngine.syncAll();
      if (!silent) {
        new Notice(
          `Entra Sync complete: ↓${summary.pulled} ↑${summary.pushed}${
            summary.deleted ? ` 🗑${summary.deleted}` : ""
          }${summary.merged > 0 ? ` (merged: ${summary.merged})` : ""}`
        );
      }
    } catch (err) {
      if (!silent) {
        new Notice(`Entra Sync failed: ${String(err)}`);
      }
    }
  }

  updateStatusBar(stateLabel = "Ready") {
    if (!this.statusBarEl) return;
    if (!this.settings.tokens) {
      this.statusBarEl.setText("Entra Sync: Not signed in");
      return;
    }
    const userLabel =
      this.settings.tokens.userDisplayName ||
      this.settings.tokens.userEmail ||
      "Entra User";
    const vaultLabel = this.settings.activeVault
      ? this.settings.activeVault.name.replace(/\s*\(.*?\)\s*$/, "")
      : "No Vault Selected";
    this.statusBarEl.setText(`Entra Sync: ${userLabel} • ${vaultLabel} (${stateLabel})`);
  }

  reinitClients() {
    const isEnterprise = this.settings.edition === "enterprise";
    const effectiveTenantId = isEnterprise
      ? this.settings.tenantId.trim()
      : "common";
    const effectiveClientId = isEnterprise
      ? this.settings.clientId.trim()
      : DEFAULT_CONSUMER_CLIENT_ID;
    const effectiveIncludeSites = isEnterprise;
    this.settings.includeSitesScope = effectiveIncludeSites;

    const graphBaseUrl =
      this.settings.graphBaseUrl || DEFAULT_GRAPH_BASE_URL;

    this.auth = new EntraPkceAuth(
      {
        tenantId: effectiveTenantId,
        clientId: effectiveClientId,
        includeSitesScope: effectiveIncludeSites,
        graphBaseUrl,
        authorityBaseUrl: this.settings.authorityBaseUrl,
      },
      this.logger
    );

    this.graph = new GraphVaultClient(
      async (forceRefresh = false) => {
        if (!this.settings.tokens) throw new Error("Not signed in");
        this.settings.tokens = await this.auth.refresh(
          this.settings.tokens,
          forceRefresh
        );
        await this.saveSettings();
        return this.settings.tokens.accessToken;
      },
      () => this.settings.graphBaseUrl || DEFAULT_GRAPH_BASE_URL,
      this.logger
    );

    if (this.settings.activeVault) {
      this.syncEngine = new CollaborativeSyncEngine(
        this.app.vault,
        this.graph,
        this.settings.activeVault,
        this.settings.syncStore,
        () => this.saveSettings(),
        (status) => this.updateStatusBar(status),
        this.logger,
        () => this.settings.syncObsidianConfig
      );
    } else {
      this.syncEngine = undefined;
    }

    if (this.syncTimerId !== null) {
      window.clearInterval(this.syncTimerId);
      this.syncTimerId = null;
    }

    if (
      this.syncEngine &&
      this.settings.tokens &&
      this.settings.autoSyncIntervalSec > 0
    ) {
      this.syncTimerId = this.registerInterval(
        window.setInterval(() => {
          void this.triggerManualSync(true);
        }, this.settings.autoSyncIntervalSec * 1000)
      );
    }

    this.updateStatusBar(this.settings.activeVault ? "Connected" : "Signed in");
  }

  async loadSettings() {
    const loaded = (await this.loadData()) as Partial<EntraSyncSettings> | null;
    this.settings = Object.assign(
      {},
      DEFAULT_SETTINGS,
      { syncStore: { files: {} } },
      loaded
    );

    // Migrate legacy pre-wizard data.json configurations smoothly
    if (!loaded?.edition) {
      const rawTenant = (loaded?.tenantId || "").trim();
      const rawClient = (loaded?.clientId || "").trim();
      const hasCustomTenant =
        Boolean(rawTenant) &&
        rawTenant.toLowerCase() !== "common" &&
        rawTenant.toLowerCase() !== "consumers";
      const hasCustomClient =
        Boolean(rawClient) && rawClient !== DEFAULT_CONSUMER_CLIENT_ID;

      if (hasCustomTenant || hasCustomClient || loaded?.includeSitesScope) {
        this.settings.edition = "enterprise";
        this.settings.tenantId = rawTenant;
        this.settings.clientId = rawClient;
      } else {
        this.settings.edition = "consumer";
        this.settings.tenantId = "";
        this.settings.clientId = "";
      }
    }

    this.settings.includeSitesScope = this.settings.edition === "enterprise";
    if (!this.settings.graphBaseUrl) {
      this.settings.graphBaseUrl = DEFAULT_GRAPH_BASE_URL;
    }
    if (!this.settings.oneDriveVaultsFolder) {
      this.settings.oneDriveVaultsFolder = "ObsidianVaults";
    }
    if (!this.settings.syncStore || !this.settings.syncStore.files) {
      this.settings.syncStore = { files: {} };
    }
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }
}

type WizardVaultMode = "discover" | "url" | "create";
type CreateVaultTarget = "sharepoint" | "onedrive";

class NonEmptyVaultConfirmationModal extends Modal {
  private resolver?: (confirmed: boolean) => void;

  constructor(app: App, private fileCount: number) {
    super(app);
  }

  confirm(): Promise<boolean> {
    return new Promise((resolve) => {
      this.resolver = resolve;
      this.open();
    });
  }

  onOpen(): void {
    this.contentEl.empty();
    this.contentEl.createEl("h2", { text: "Local vault is not empty" });
    this.contentEl.createEl("p", {
      text: `This local vault already contains ${this.fileCount} ${
        this.fileCount === 1 ? "file" : "files"
      }. Connecting to an existing remote vault will upload those files.`,
    });
    this.contentEl.createEl("p", {
      text: "Are you sure you want to continue? To avoid uploading them, cancel and create a new vault or use an empty local folder.",
    });

    new Setting(this.contentEl)
      .addButton((button) =>
        button.setButtonText("No, cancel").onClick(() => {
          this.finish(false);
        })
      )
      .addButton((button) =>
        button
          .setButtonText("Yes, upload existing files")
          .setCta()
          .onClick(() => {
            this.finish(true);
          })
      );
  }

  onClose(): void {
    this.finish(false, false);
  }

  private finish(confirmed: boolean, close = true): void {
    const resolver = this.resolver;
    this.resolver = undefined;
    resolver?.(confirmed);
    if (close) {
      this.close();
    }
  }
}

class EntraSyncSettingTab extends PluginSettingTab {
  private wizardStep: 1 | 2 | 3 = 1;
  private switchingVault = false;
  private vaultMode: WizardVaultMode = "discover";
  private createTarget: CreateVaultTarget = "sharepoint";
  private createParentUrl = "";
  private availableVaults: RemoteVaultTarget[] = [];
  private discoverLoading = false;
  private discoveredKey = "";
  private editingOneDriveFolder = false;
  private pendingOneDriveFolder = "";
  private newVaultName = "";
  private vaultUrlInput = "";
  private enterpriseEmail = "";
  private directoryLookupSucceeded = false;

  constructor(app: App, private plugin: EntraSharePointSyncPlugin) {
    super(app, plugin);
    this.createTarget =
      this.plugin.settings.edition === "enterprise" ? "sharepoint" : "onedrive";
  }

  onAuthCompleted(): void {
    this.wizardStep = 3;
    this.discoveredKey = "";
    this.createTarget =
      this.plugin.settings.edition === "enterprise" ? "sharepoint" : "onedrive";
    this.display();
  }

  private async connectToVault(
    target: RemoteVaultTarget,
    confirmNonEmptyVault = true
  ): Promise<void> {
    if (!(await this.plugin.graph.isObsidianVault(target))) {
      throw new Error(
        "This is not an existing Obsidian vault. Use Create vault to initialize this folder first."
      );
    }

    const localContentCount = this.app.vault
      .getFiles()
      .filter(
        (file) =>
          !file.path.startsWith(`${this.app.vault.configDir}/`) &&
          !isIgnoredVaultPath(
            file.path,
            this.plugin.settings.syncObsidianConfig,
            this.app.vault.configDir
          )
      ).length;

    if (
      confirmNonEmptyVault &&
      localContentCount > 0 &&
      !(await new NonEmptyVaultConfirmationModal(
        this.app,
        localContentCount
      ).confirm())
    ) {
      return;
    }

    const previousVault = this.plugin.settings.activeVault;
    const previousStore = this.plugin.settings.syncStore;
    this.plugin.settings.activeVault = target;
    this.plugin.settings.syncStore = { files: {} };
    await this.plugin.saveSettings();
    this.plugin.reinitClients();
    try {
      if (!this.plugin.syncEngine) {
        throw new Error("Sync engine could not be initialized.");
      }
      await this.plugin.syncEngine.syncAll();
      this.switchingVault = false;
      await this.plugin.saveSettings();
      new Notice(`Connected to vault: ${target.name}`);
      this.display();
    } catch (err) {
      this.plugin.settings.activeVault = previousVault;
      this.plugin.settings.syncStore = previousStore;
      await this.plugin.saveSettings();
      this.plugin.reinitClients();
      throw err;
    }
  }

  private async refreshDiscoveredVaults(showNotice = false): Promise<void> {
    if (!this.plugin.settings.tokens || this.discoverLoading) {
      return;
    }
    this.discoverLoading = true;
    const currentKey = `${this.plugin.settings.edition}:${
      this.plugin.settings.tokens.userId ||
      this.plugin.settings.tokens.userEmail ||
      "signed-in"
    }:${this.plugin.settings.oneDriveVaultsFolder}`;

    try {
      this.availableVaults = await this.plugin.graph.listAvailableVaults(
        this.plugin.auth.isSharePointEnabled(),
        this.plugin.settings.oneDriveVaultsFolder
      );
      this.discoveredKey = currentKey;
      if (showNotice) {
        new Notice(`Found ${this.availableVaults.length} vault(s).`);
      }
    } catch (err) {
      this.discoveredKey = currentKey;
      new Notice(`Failed to list vaults: ${String(err)}`);
    } finally {
      this.discoverLoading = false;
      this.display();
    }
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    const isSignedIn = Boolean(this.plugin.settings.tokens);
    const hasActiveVault = Boolean(this.plugin.settings.activeVault);

    if (isSignedIn && hasActiveVault && !this.switchingVault) {
      this.renderConnectedView(containerEl);
      this.renderTroubleshootingSection(containerEl);
      return;
    }

    if (isSignedIn) {
      this.wizardStep = 3;
    } else if (this.wizardStep === 3) {
      this.wizardStep = 1;
    }

    this.renderStepper(containerEl);

    if (this.wizardStep === 1) {
      this.renderStep1Edition(containerEl);
    } else if (this.wizardStep === 2) {
      this.renderStep2SignIn(containerEl);
    } else {
      this.renderStep3ConnectVault(containerEl);
    }

    this.renderTroubleshootingSection(containerEl);
  }

  private renderStepper(containerEl: HTMLElement): void {
    const stepper = containerEl.createDiv({ cls: "entra-sync-stepper" });
    const steps: Array<{ num: 1 | 2 | 3; label: string }> = [
      { num: 1, label: "Account type" },
      { num: 2, label: "Sign in" },
      { num: 3, label: "Connect vault" },
    ];

    steps.forEach((step, idx) => {
      const isDone = this.wizardStep > step.num;
      const isActive = this.wizardStep === step.num;
      const stepEl = stepper.createDiv({
        cls: `entra-sync-step${isActive ? " is-active" : ""}${
          isDone ? " is-done" : ""
        }`,
      });
      stepEl.createEl("span", {
        cls: "entra-sync-step-num",
        text: isDone ? "✓" : String(step.num),
      });
      stepEl.createEl("span", { text: step.label });

      if (isDone && !this.plugin.settings.tokens) {
        stepEl.addEventListener("click", () => {
          this.wizardStep = step.num;
          this.display();
        });
      }

      if (idx < steps.length - 1) {
        stepper.createDiv({ cls: "entra-sync-step-line" });
      }
    });
  }

  private renderStep1Edition(containerEl: HTMLElement): void {
    new Setting(containerEl)
      .setName("Choose your Microsoft account type")
      .setDesc(
        "Select Consumer for personal OneDrive accounts, or Enterprise for corporate SharePoint Online, Microsoft Teams, and OneDrive for Business."
      )
      .setHeading();

    const grid = containerEl.createDiv({ cls: "entra-sync-choice-grid" });

    const selectEdition = async (edition: AccountEdition) => {
      this.plugin.settings.edition = edition;
      this.plugin.settings.includeSitesScope = edition === "enterprise";
      this.createTarget = edition === "enterprise" ? "sharepoint" : "onedrive";
      await this.plugin.saveSettings();
      this.plugin.reinitClients();
      this.display();
    };

    const consumerCard = grid.createDiv({
      cls: `entra-sync-choice-card${
        this.plugin.settings.edition === "consumer" ? " is-selected" : ""
      }`,
    });
    consumerCard.createDiv({
      cls: "entra-sync-choice-title",
      text: "Consumer (OneDrive)",
    });
    consumerCard.createDiv({
      cls: "entra-sync-choice-desc",
      text: "Use a personal Microsoft account (outlook.com, live.com, hotmail.com) to sync personal or shared OneDrive folders.",
    });
    consumerCard.createDiv({
      cls: "entra-sync-choice-meta",
      text: "No admin setup required",
    });
    consumerCard.addEventListener("click", () => {
      void selectEdition("consumer");
    });

    const enterpriseCard = grid.createDiv({
      cls: `entra-sync-choice-card${
        this.plugin.settings.edition === "enterprise" ? " is-selected" : ""
      }`,
    });
    enterpriseCard.createDiv({
      cls: "entra-sync-choice-title",
      text: "Enterprise (Teams & SharePoint)",
    });
    enterpriseCard.createDiv({
      cls: "entra-sync-choice-desc",
      text: "Use your corporate Microsoft Entra ID account to collaborate in SharePoint Online, Microsoft Teams channels, and OneDrive for Business.",
    });
    enterpriseCard.createDiv({
      cls: "entra-sync-choice-meta",
      text: "Requires your organization's Entra app registration IDs",
    });
    enterpriseCard.addEventListener("click", () => {
      void selectEdition("enterprise");
    });

    const nav = containerEl.createDiv({ cls: "entra-sync-wizard-nav" });
    nav.createDiv();
    new Setting(nav).addButton((btn) =>
      btn
        .setButtonText("Continue")
        .setCta()
        .onClick(() => {
          this.wizardStep = 2;
          this.display();
        })
    );
  }

  private renderStep2SignIn(containerEl: HTMLElement): void {
    if (this.plugin.settings.edition === "consumer") {
      new Setting(containerEl)
        .setName("Sign in with your personal Microsoft account")
        .setDesc(
          "Opens your browser using OAuth 2.0 PKCE and returns to Obsidian automatically. No credentials are entered inside Obsidian."
        )
        .setHeading();

      new Setting(containerEl)
        .setName("Personal Microsoft OneDrive")
        .setDesc(
          "Requests User.Read, Files.ReadWrite.All, and offline_access permissions."
        )
        .addButton((btn) =>
          btn
            .setButtonText("Sign in")
            .setCta()
            .onClick(async () => {
              this.plugin.reinitClients();
              const url = await this.plugin.auth.startLogin();
              openExternalUrl(url);
            })
        );

      const nav = containerEl.createDiv({ cls: "entra-sync-wizard-nav" });
      new Setting(nav).addButton((btn) =>
        btn.setButtonText("Back").onClick(() => {
          this.wizardStep = 1;
          this.display();
        })
      );
      return;
    }

    // Enterprise Step 2
    new Setting(containerEl)
      .setName("Enterprise accounts require configuration from your admin")
      .setDesc(
        "Use your organization's configured Microsoft Entra ID app registration."
      )
      .setHeading();

    const adminDetailsEl = containerEl.createEl("details", {
      cls: "entra-sync-details",
    });
    adminDetailsEl.createEl("summary", {
      text: "View one-time IT admin setup request",
    });
    const detailsBody = adminDetailsEl.createDiv({
      cls: "entra-sync-details-body",
    });
    detailsBody.createEl("pre", {
      cls: "entra-sync-version-preview",
      text: ADMIN_REQUEST_TEMPLATE,
    });
    new Setting(detailsBody).addButton((btn) =>
      btn.setButtonText("Copy admin request").onClick(async () => {
        await navigator.clipboard.writeText(ADMIN_REQUEST_TEMPLATE);
        new Notice("Copied IT admin setup request to clipboard.");
      })
    );

    new Setting(containerEl)
      .setName("Enter your email to check the configuration")
      .setDesc(
        "If your organization has been added to the plugin, its Directory and Application IDs are filled in automatically."
      )
      .addText((text) =>
        text
          .setPlaceholder("you@company.com")
          .setValue(this.enterpriseEmail)
          .onChange((val) => {
            this.enterpriseEmail = val;
            this.directoryLookupSucceeded = false;
          })
      )
      .addButton((btn) =>
        btn.setButtonText("Check configuration").onClick(async () => {
          const configuration = findEnterpriseDirectoryConfiguration(
            this.enterpriseEmail
          );
          if (!configuration) {
            this.directoryLookupSucceeded = false;
            new Notice(
              "No built-in configuration found for that email. Ask your admin for the IDs or enter the configuration manually."
            );
            this.display();
            return;
          }

          this.plugin.settings.tenantId = configuration.tenantId;
          this.plugin.settings.clientId = configuration.clientId;
          this.directoryLookupSucceeded = true;
          await this.plugin.saveSettings();
          this.plugin.reinitClients();
          new Notice(
            `Configuration found for ${configuration.emailDomain}. You can sign in now.`
          );
          this.display();
        })
      );

    if (this.directoryLookupSucceeded) {
      containerEl.createDiv({
        cls: "entra-sync-lookup-status",
        text: "Configuration found. Your Directory and Application IDs are ready.",
      });
    }

    const manualDetailsEl = containerEl.createEl("details", {
      cls: "entra-sync-details",
    });
    manualDetailsEl.createEl("summary", {
      text: "Enter configuration manually",
    });
    const manualBody = manualDetailsEl.createDiv({
      cls: "entra-sync-details-body",
    });

    new Setting(manualBody)
      .setName("Directory (tenant) ID")
      .setDesc(
        "Your organization's Microsoft Entra ID tenant UUID or domain (for example contoso.onmicrosoft.com)."
      )
      .addText((text) =>
        text
          .setPlaceholder("00000000-0000-0000-0000-000000000000")
          .setValue(this.plugin.settings.tenantId)
          .onChange(async (val) => {
            this.plugin.settings.tenantId = val.trim();
            await this.plugin.saveSettings();
            this.plugin.reinitClients();
          })
      );

    new Setting(manualBody)
      .setName("Application (client) ID")
      .setDesc(
        "Application (client) ID from your organization's Microsoft Entra app registration."
      )
      .addText((text) =>
        text
          .setPlaceholder("00000000-0000-0000-0000-000000000000")
          .setValue(this.plugin.settings.clientId)
          .onChange(async (val) => {
            this.plugin.settings.clientId = val.trim();
            await this.plugin.saveSettings();
            this.plugin.reinitClients();
          })
      );

    const nav = containerEl.createDiv({ cls: "entra-sync-wizard-nav" });
    new Setting(nav)
      .addButton((btn) =>
        btn.setButtonText("Back").onClick(() => {
          this.wizardStep = 1;
          this.display();
        })
      )
      .addButton((btn) =>
        btn
          .setButtonText("Sign in")
          .setCta()
          .onClick(async () => {
            if (
              !this.plugin.settings.tenantId.trim() ||
              !this.plugin.settings.clientId.trim()
            ) {
              new Notice(
                "Enter both Directory (tenant) ID and Application (client) ID first."
              );
              return;
            }
            this.plugin.reinitClients();
            const url = await this.plugin.auth.startLogin();
            openExternalUrl(url);
          })
      );
  }

  private renderStep3ConnectVault(containerEl: HTMLElement): void {
    const tokens = this.plugin.settings.tokens;
    const isEnterprise = this.plugin.settings.edition === "enterprise";
    const userDisplay =
      tokens?.userDisplayName || tokens?.userEmail || "Authenticated user";
    const userEmail = tokens?.userEmail || "";

    const accountCard = containerEl.createDiv({ cls: "entra-sync-card" });
    const header = accountCard.createDiv({ cls: "entra-sync-card-header" });
    const info = header.createDiv();
    info.createDiv({
      cls: "entra-sync-badge",
      text: `✓ Signed in as ${userDisplay}${
        userEmail && userEmail !== userDisplay ? ` (${userEmail})` : ""
      }`,
    });
    info.createEl("p", {
      cls: "entra-sync-muted",
      text: isEnterprise
        ? `Enterprise (Teams & SharePoint) • Tenant: ${this.plugin.settings.tenantId}`
        : "Consumer (OneDrive)",
    });

    const actions = header.createDiv({ cls: "entra-sync-card-actions" });
    const cardActionsSetting = new Setting(actions);
    if (this.switchingVault && this.plugin.settings.activeVault) {
      cardActionsSetting.addButton((btn) =>
        btn.setButtonText("Cancel").onClick(() => {
          this.switchingVault = false;
          this.display();
        })
      );
    }
    cardActionsSetting.addButton((btn) =>
      btn.setButtonText("Change account").onClick(async () => {
        this.plugin.settings.tokens = undefined;
        this.plugin.settings.activeVault = undefined;
        this.plugin.settings.syncStore = { files: {} };
        this.switchingVault = false;
        this.wizardStep = 1;
        this.availableVaults = [];
        this.discoveredKey = "";
        await this.plugin.saveSettings();
        this.plugin.reinitClients();
        this.display();
      })
    );

    const tabs = containerEl.createDiv({ cls: "entra-sync-action-tabs" });
    const tabDefs: Array<{
      mode: WizardVaultMode;
      title: string;
      desc: string;
    }> = [
      {
        mode: "discover",
        title: "Discover vaults",
        desc: isEnterprise
          ? "Pick from followed SharePoint & Teams sites, shared folders, or OneDrive"
          : "Pick from your OneDrive or folders shared with you",
      },
      {
        mode: "url",
        title: "Open vault by URL",
        desc: isEnterprise
          ? "Paste a SharePoint, Microsoft Teams, or OneDrive folder link"
          : "Paste a OneDrive folder URL or sharing link",
      },
      {
        mode: "create",
        title: "Create vault",
        desc: isEnterprise
          ? "Create a new vault folder in SharePoint, Teams, or OneDrive"
          : "Create a new vault folder in your OneDrive",
      },
    ];

    for (const tabDef of tabDefs) {
      const tabEl = tabs.createDiv({
        cls: `entra-sync-action-tab${
          this.vaultMode === tabDef.mode ? " is-active" : ""
        }`,
      });
      tabEl.createDiv({ cls: "entra-sync-action-tab-title", text: tabDef.title });
      tabEl.createDiv({ cls: "entra-sync-action-tab-desc", text: tabDef.desc });
      tabEl.addEventListener("click", () => {
        this.vaultMode = tabDef.mode;
        this.display();
      });
    }

    if (this.vaultMode === "discover") {
      this.renderDiscoverPanel(containerEl, isEnterprise);
    } else if (this.vaultMode === "url") {
      this.renderOpenUrlPanel(containerEl, isEnterprise);
    } else {
      this.renderCreateVaultPanel(containerEl, isEnterprise);
    }
  }

  private renderDiscoverPanel(
    containerEl: HTMLElement,
    isEnterprise: boolean
  ): void {
    const panel = containerEl.createDiv({ cls: "entra-sync-panel" });
    const tokens = this.plugin.settings.tokens;
    const expectedKey = `${this.plugin.settings.edition}:${
      tokens?.userId || tokens?.userEmail || "signed-in"
    }:${this.plugin.settings.oneDriveVaultsFolder}`;

    if (this.discoveredKey !== expectedKey && !this.discoverLoading) {
      void this.refreshDiscoveredVaults(false);
    }

    new Setting(panel)
      .setName("Available vaults")
      .setDesc(
        isEnterprise
          ? "Scans followed SharePoint & Teams sites, folders shared with you, and your OneDrive for folders containing an existing .obsidian directory."
          : "Scans your OneDrive vaults folder and folders shared with you for folders containing an existing .obsidian directory."
      )
      .addButton((btn) =>
        btn
          .setButtonText(this.discoverLoading ? "Scanning..." : "Refresh")
          .onClick(async () => {
            await this.refreshDiscoveredVaults(true);
          })
      );

    const listEl = panel.createDiv({ cls: "entra-sync-vault-list" });
    if (this.discoverLoading) {
      listEl.createEl("p", {
        cls: "entra-sync-muted",
        text: "Scanning Microsoft Graph for available vaults...",
      });
    } else if (this.availableVaults.length === 0) {
      listEl.createEl("p", {
        cls: "entra-sync-muted",
        text: "No vaults found yet. Open a vault by URL or create a new vault.",
      });
    } else {
      for (const vault of this.availableVaults) {
        const match = vault.name.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
        const cleanTitle = match ? match[1] : vault.name;
        const badgeLabel = match ? match[2] : "Remote vault";

        const row = listEl.createDiv({ cls: "entra-sync-vault-row" });
        const info = row.createDiv({ cls: "entra-sync-vault-info" });
        const titleRow = info.createDiv({ cls: "entra-sync-vault-title" });
        titleRow.createEl("span", { text: cleanTitle });
        titleRow.createEl("span", {
          cls: "entra-sync-source-badge",
          text: badgeLabel,
        });
        if (vault.webUrl) {
          info.createDiv({ cls: "entra-sync-vault-meta", text: vault.webUrl });
        }

        new Setting(row).addButton((btn) =>
          btn
            .setButtonText("Connect")
            .setCta()
            .onClick(async () => {
              try {
                await this.connectToVault(vault);
              } catch (err) {
                new Notice(`Failed to connect to vault: ${String(err)}`);
              }
            })
        );
      }
    }

    // Inline OneDrive parent folder control (applicable to OneDrive scanning only)
    const folderBar = panel.createDiv({ cls: "entra-sync-inline-folder" });
    if (!this.editingOneDriveFolder) {
      folderBar.createEl("span", {
        text: `OneDrive parent folder: /${this.plugin.settings.oneDriveVaultsFolder}`,
      });
      new Setting(folderBar).addButton((btn) =>
        btn.setButtonText("Change folder...").onClick(() => {
          this.pendingOneDriveFolder = this.plugin.settings.oneDriveVaultsFolder;
          this.editingOneDriveFolder = true;
          this.display();
        })
      );
    } else {
      new Setting(folderBar)
        .setName("OneDrive vaults folder")
        .addText((text) =>
          text
            .setPlaceholder("ObsidianVaults")
            .setValue(this.pendingOneDriveFolder)
            .onChange((val) => {
              this.pendingOneDriveFolder = val;
            })
        )
        .addButton((btn) =>
          btn
            .setButtonText("Save and rescan")
            .setCta()
            .onClick(async () => {
              this.plugin.settings.oneDriveVaultsFolder =
                this.pendingOneDriveFolder.trim() || "ObsidianVaults";
              this.editingOneDriveFolder = false;
              this.discoveredKey = "";
              await this.plugin.saveSettings();
              await this.refreshDiscoveredVaults(true);
            })
        )
        .addButton((btn) =>
          btn.setButtonText("Cancel").onClick(() => {
            this.editingOneDriveFolder = false;
            this.display();
          })
        );
    }
  }

  private renderOpenUrlPanel(
    containerEl: HTMLElement,
    isEnterprise: boolean
  ): void {
    const panel = containerEl.createDiv({ cls: "entra-sync-panel" });
    new Setting(panel)
      .setName("Vault folder URL")
      .setDesc(
        isEnterprise
          ? "Paste a SharePoint library or folder URL, a Microsoft Teams channel URL (in Teams: Files -> Open in SharePoint -> copy browser URL), or a sharing link. The folder must already contain .obsidian."
          : "Paste a shared OneDrive folder link or OneDrive folder URL from your browser. The folder must already contain .obsidian."
      )
      .addText((text) =>
        text
          .setPlaceholder(
            isEnterprise
              ? "https://contoso.sharepoint.com/sites/Engineering/Shared%20Documents/Vault"
              : "https://onedrive.live.com/... or https://1drv.ms/f/..."
          )
          .setValue(this.vaultUrlInput)
          .onChange((val) => {
            this.vaultUrlInput = val;
          })
      )
      .addButton((btn) =>
        btn
          .setButtonText("Open vault")
          .setCta()
          .onClick(async () => {
            if (!this.vaultUrlInput.trim()) {
              new Notice("Paste a vault folder URL first.");
              return;
            }
            try {
              const resolved = await this.plugin.graph.resolveVaultFromUrl(
                this.vaultUrlInput.trim()
              );
              await this.connectToVault(resolved);
            } catch (err) {
              new Notice(`Failed to open vault URL: ${String(err)}`);
            }
          })
      );
  }

  private renderCreateVaultPanel(
    containerEl: HTMLElement,
    isEnterprise: boolean
  ): void {
    const panel = containerEl.createDiv({ cls: "entra-sync-panel" });

    if (isEnterprise) {
      const pills = panel.createDiv({ cls: "entra-sync-target-pills" });
      const spPill = pills.createDiv({
        cls: `entra-sync-target-pill${
          this.createTarget === "sharepoint" ? " is-active" : ""
        }`,
        text: "SharePoint or Teams library",
      });
      spPill.addEventListener("click", () => {
        this.createTarget = "sharepoint";
        this.display();
      });

      const odPill = pills.createDiv({
        cls: `entra-sync-target-pill${
          this.createTarget === "onedrive" ? " is-active" : ""
        }`,
        text: "My OneDrive for Business",
      });
      odPill.addEventListener("click", () => {
        this.createTarget = "onedrive";
        this.display();
      });
    }

    new Setting(panel)
      .setName("Vault folder name")
      .setDesc(
        "Name of the remote vault. If this folder already exists, it is initialized in place so its files can merge with the local vault."
      )
      .addText((text) =>
        text
          .setPlaceholder("e.g. ProductArchitecture")
          .setValue(this.newVaultName)
          .onChange((val) => {
            this.newVaultName = val;
          })
      );

    const useSharePointParent = isEnterprise && this.createTarget === "sharepoint";

    if (useSharePointParent) {
      new Setting(panel)
        .setName("Parent SharePoint or Teams library URL")
        .setDesc(
          "Paste the SharePoint Document Library or parent folder URL where the new vault folder should be created."
        )
        .addText((text) =>
          text
            .setPlaceholder(
              "https://contoso.sharepoint.com/sites/Engineering/Shared%20Documents"
            )
            .setValue(this.createParentUrl)
            .onChange((val) => {
              this.createParentUrl = val;
            })
        );
    } else {
      new Setting(panel)
        .setName("Parent folder in OneDrive")
        .setDesc(
          `Creates /${this.plugin.settings.oneDriveVaultsFolder}/${
            this.newVaultName.trim() || "MyVault"
          } in your OneDrive.`
        )
        .addText((text) =>
          text
            .setPlaceholder("ObsidianVaults")
            .setValue(this.plugin.settings.oneDriveVaultsFolder)
            .onChange(async (val) => {
              this.plugin.settings.oneDriveVaultsFolder =
                val.trim() || "ObsidianVaults";
              await this.plugin.saveSettings();
            })
        );
    }

    new Setting(panel).addButton((btn) =>
      btn
        .setButtonText("Create and connect vault")
        .setCta()
        .onClick(async () => {
          if (!this.newVaultName.trim()) {
            new Notice("Enter a vault folder name first.");
            return;
          }
          if (useSharePointParent && !this.createParentUrl.trim()) {
            new Notice(
              "Paste the parent SharePoint or Microsoft Teams library URL first."
            );
            return;
          }
          try {
            const parentArg = useSharePointParent
              ? this.createParentUrl.trim()
              : this.plugin.settings.oneDriveVaultsFolder;
            const created = await this.plugin.graph.createOrGetVaultFolder(
              this.newVaultName.trim(),
              parentArg
            );
            await this.connectToVault(created, false);
          } catch (err) {
            new Notice(`Failed to create vault: ${String(err)}`);
          }
        })
    );
  }

  private renderConnectedView(containerEl: HTMLElement): void {
    const activeVault = this.plugin.settings.activeVault!;
    const tokens = this.plugin.settings.tokens!;
    const isEnterprise = this.plugin.settings.edition === "enterprise";

    // Connected Vault Card
    const vaultCard = containerEl.createDiv({ cls: "entra-sync-card" });
    const vaultHeader = vaultCard.createDiv({ cls: "entra-sync-card-header" });
    const vaultInfo = vaultHeader.createDiv();
    vaultInfo.createDiv({
      cls: "entra-sync-badge",
      text: `✓ Connected vault: ${activeVault.name}`,
    });
    if (activeVault.webUrl) {
      vaultInfo.createEl("p", {
        cls: "entra-sync-muted",
        text: activeVault.webUrl,
      });
    }

    const vaultActions = vaultHeader.createDiv({ cls: "entra-sync-card-actions" });
    new Setting(vaultActions)
      .addButton((btn) =>
        btn
          .setButtonText("Sync now")
          .setCta()
          .onClick(async () => {
            await this.plugin.triggerManualSync();
          })
      )
      .addButton((btn) =>
        btn.setButtonText("Share vault...").onClick(() => {
          if (this.plugin.settings.activeVault) {
            new ShareVaultModal(
              this.app,
              this.plugin.settings.activeVault,
              this.plugin.graph
            ).open();
          }
        })
      )
      .addButton((btn) =>
        btn.setButtonText("Switch vault...").onClick(() => {
          this.switchingVault = true;
          this.wizardStep = 3;
          this.discoveredKey = "";
          this.display();
        })
      );

    // Account Card
    const userDisplay =
      tokens.userDisplayName || tokens.userEmail || "Authenticated user";
    const userEmail = tokens.userEmail || "";
    const accountCard = containerEl.createDiv({ cls: "entra-sync-card" });
    const accountHeader = accountCard.createDiv({ cls: "entra-sync-card-header" });
    const accountInfo = accountHeader.createDiv();
    accountInfo.createEl("div", {
      cls: "entra-sync-choice-title",
      text: `${userDisplay}${
        userEmail && userEmail !== userDisplay ? ` (${userEmail})` : ""
      }`,
    });
    accountInfo.createEl("p", {
      cls: "entra-sync-muted",
      text: isEnterprise
        ? `Enterprise (Teams & SharePoint) • Tenant: ${this.plugin.settings.tenantId}`
        : "Consumer (OneDrive)",
    });

    const accountActions = accountHeader.createDiv({
      cls: "entra-sync-card-actions",
    });
    new Setting(accountActions).addButton((btn) =>
      btn.setButtonText("Sign out").onClick(async () => {
        this.plugin.settings.tokens = undefined;
        this.plugin.settings.activeVault = undefined;
        this.plugin.settings.syncStore = { files: {} };
        this.switchingVault = false;
        this.wizardStep = 1;
        this.availableVaults = [];
        this.discoveredKey = "";
        await this.plugin.saveSettings();
        this.plugin.reinitClients();
        this.display();
      })
    );

    // Sync Preferences
    new Setting(containerEl).setName("Sync preferences").setHeading();

    const configDir = this.app.vault.configDir;
    new Setting(containerEl)
      .setName(`Sync vault configuration (${configDir})`)
      .setDesc(
        `Sync shared vault configuration, themes, CSS snippets, hotkeys, and plugins inside ${configDir}/ (excludes your personal OAuth tokens in data.json and personal workspace.json tab layout).`
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.syncObsidianConfig)
          .onChange(async (val) => {
            this.plugin.settings.syncObsidianConfig = val;
            await this.plugin.saveSettings();
            if (val) {
              await this.plugin.triggerManualSync();
            }
          })
      );

    new Setting(containerEl)
      .setName("Note version history")
      .setDesc(
        "Browse and restore previous SharePoint or OneDrive revisions of the active note."
      )
      .addButton((btn) =>
        btn.setButtonText("Active note version history").onClick(() => {
          this.plugin.openActiveFileHistory();
        })
      );
  }

  private renderTroubleshootingSection(containerEl: HTMLElement): void {
    const detailsEl = containerEl.createEl("details", {
      cls: "entra-sync-details",
    });
    detailsEl.createEl("summary", {
      text: "Troubleshooting and diagnostic logs",
    });
    const body = detailsEl.createDiv({ cls: "entra-sync-details-body" });

    if (this.plugin.settings.tokens) {
      const expDate = new Date(
        this.plugin.settings.tokens.expiresAt
      ).toLocaleString();
      const hasRefresh = Boolean(this.plugin.settings.tokens.refreshToken);
      new Setting(body)
        .setName("OAuth2 token and delta diagnostics")
        .setDesc(
          `Access token expires: ${expDate} • Refresh token present: ${
            hasRefresh ? "Yes" : "No"
          }`
        )
        .addButton((btn) =>
          btn.setButtonText("Force token refresh").onClick(async () => {
            if (!this.plugin.settings.tokens) return;
            try {
              this.plugin.settings.tokens = await this.plugin.auth.refresh(
                this.plugin.settings.tokens,
                true
              );
              await this.plugin.saveSettings();
              new Notice("Access token refreshed successfully!");
              this.display();
            } catch (err) {
              new Notice(`Token refresh failed: ${String(err)}`);
              this.display();
            }
          })
        )
        .addButton((btn) =>
          btn.setButtonText("Reset delta and resync").onClick(async () => {
            this.plugin.settings.syncStore.deltaLink = undefined;
            await this.plugin.saveSettings();
            this.plugin.logger.info(
              "Cleared cached deltaLink; triggering full delta sync"
            );
            await this.plugin.triggerManualSync();
            this.display();
          })
        );
    }

    const logPre = body.createEl("pre", {
      cls: "entra-sync-version-preview entra-sync-log-viewer",
      text: this.plugin.logger.formatText(),
    });

    new Setting(body)
      .setName("Diagnostic log actions")
      .setDesc(
        "Warnings and errors are also logged to the developer console with prefix [EntraSync]."
      )
      .addButton((btn) =>
        btn.setButtonText("Refresh logs").onClick(() => {
          logPre.setText(this.plugin.logger.formatText());
        })
      )
      .addButton((btn) =>
        btn.setButtonText("Copy logs").onClick(async () => {
          await navigator.clipboard.writeText(this.plugin.logger.formatText());
          new Notice("Copied diagnostic logs to clipboard.");
        })
      )
      .addButton((btn) =>
        btn.setButtonText("Clear logs").onClick(() => {
          this.plugin.logger.clear();
          logPre.setText(this.plugin.logger.formatText());
        })
      );
  }
}

class SyncLogsModal extends Modal {
  constructor(app: App, private plugin: EntraSharePointSyncPlugin) {
    super(app);
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h2", { text: "Sync diagnostic logs" });
    const pre = contentEl.createEl("pre", {
      cls: "entra-sync-version-preview entra-sync-log-viewer",
      text: this.plugin.logger.formatText(),
    });

    new Setting(contentEl)
      .addButton((btn) =>
        btn.setButtonText("Refresh").onClick(() => {
          pre.setText(this.plugin.logger.formatText());
        })
      )
      .addButton((btn) =>
        btn
          .setButtonText("Copy to clipboard")
          .setCta()
          .onClick(async () => {
            await navigator.clipboard.writeText(this.plugin.logger.formatText());
            new Notice("Copied diagnostic logs to clipboard.");
          })
      );
  }
}
