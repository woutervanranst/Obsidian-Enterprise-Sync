import { App, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile, TFolder } from "obsidian";
import { AuthTokens, EntraPkceAuth } from "./auth";
import { GraphVaultClient, RemoteVaultTarget } from "./graph";
import { CollaborativeSyncEngine, SyncStore, isIgnoredVaultPath } from "./sync";
import { VersionHistoryModal } from "./historyModal";
import { ShareVaultModal } from "./shareModal";
import { SyncLogger, defaultLogger } from "./logger";

export interface EntraSyncSettings {
  tenantId: string;
  clientId: string;
  graphBaseUrl: string;
  oneDriveVaultsFolder: string;
  includeSitesScope: boolean;
  autoSyncIntervalSec: number;
  tokens?: AuthTokens;
  activeVault?: RemoteVaultTarget;
  syncStore: SyncStore;
}

const DEFAULT_SETTINGS: EntraSyncSettings = {
  tenantId: "common", // Works for Consumer OneDrive + Multitenant Entra ID
  clientId: "ada87c12-a404-484b-a7dc-10c54beb2d7d",
  graphBaseUrl: "https://graph.microsoft.com/v1.0",
  oneDriveVaultsFolder: "ObsidianVaults",
  includeSitesScope: false,
  autoSyncIntervalSec: 15,
  syncStore: { files: {} },
};

export default class EntraSharePointSyncPlugin extends Plugin {
  settings: EntraSyncSettings = DEFAULT_SETTINGS;
  logger: SyncLogger = defaultLogger;
  auth!: EntraPkceAuth;
  graph!: GraphVaultClient;
  syncEngine?: CollaborativeSyncEngine;
  statusBarEl?: HTMLElement;
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
          new Notice(
            `Signed in with Microsoft as ${
              this.settings.tokens.userDisplayName ||
              this.settings.tokens.userEmail ||
              "Enterprise User"
            }!`
          );
        } catch (e) {
          new Notice(`Sign-in failed: ${String(e)}`);
        }
      }
    });

    // Ribbon: Sync Now
    this.addRibbonIcon("refresh-cw", "Entra Sync: Sync Vault Now", async () => {
      await this.triggerManualSync();
    });

    // Ribbon: Share Vault
    this.addRibbonIcon("users", "Entra Sync: Share Vault with Colleagues", () => {
      if (!this.settings.tokens) {
        new Notice("Please sign in with Entra ID / Microsoft first.");
        return;
      }
      if (!this.settings.activeVault) {
        new Notice("Please select or create a shared vault in settings first.");
        return;
      }
      new ShareVaultModal(this.app, this.settings.activeVault, this.graph).open();
    });

    // Ribbon: Version History
    this.addRibbonIcon("history", "Entra Sync: View Note Version History", () => {
      this.openActiveFileHistory();
    });

    // Auto-push when user creates or modifies a note or attachment (debounced)
    const schedulePush = (abstractFile: unknown) => {
      if (
        this.syncEngine &&
        abstractFile instanceof TFile &&
        !isIgnoredVaultPath(abstractFile.path) &&
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
            this.updateStatusBar("Sync Error");
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
          !isIgnoredVaultPath(abstractFile.path) &&
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
              this.updateStatusBar("Sync Error");
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
          !isIgnoredVaultPath(abstractFile.path)
        ) {
          void this.syncEngine
            .renameLocalPath(oldPath, abstractFile)
            .then(() => this.updateStatusBar("Synced"))
            .catch((err) => {
              this.logger.error(
                `Remote rename failed from ${oldPath} to ${abstractFile.path}`,
                String(err)
              );
              this.updateStatusBar("Sync Error");
            });
        }
      })
    );

    // Command: Sync Vault Now
    this.addCommand({
      id: "entra-sync-now",
      name: "Sync Vault Now (SharePoint / OneDrive)",
      callback: async () => {
        await this.triggerManualSync();
      },
    });

    // Command: Share Vault with Colleagues
    this.addCommand({
      id: "entra-share-vault",
      name: "Share Vault with Colleagues (SharePoint / OneDrive)",
      callback: () => {
        if (!this.settings.activeVault) {
          new Notice("No active shared vault configured.");
          return;
        }
        new ShareVaultModal(this.app, this.settings.activeVault, this.graph).open();
      },
    });

    // Command: View SharePoint / OneDrive Version History for active note
    this.addCommand({
      id: "view-note-version-history",
      name: "View Note Version History (SharePoint / OneDrive)",
      callback: () => {
        this.openActiveFileHistory();
      },
    });

    // Command: View Sync Diagnostic Logs
    this.addCommand({
      id: "entra-view-sync-logs",
      name: "View Sync Diagnostic Logs (SharePoint / OneDrive)",
      callback: () => {
        new SyncLogsModal(this.app, this).open();
      },
    });

    this.addSettingTab(new EntraSyncSettingTab(this.app, this));

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
    this.auth = new EntraPkceAuth(
      {
        tenantId: this.settings.tenantId,
        clientId: this.settings.clientId,
        includeSitesScope: this.settings.includeSitesScope,
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
      () => this.settings.graphBaseUrl || "https://graph.microsoft.com/v1.0",
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
        this.logger
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
      this.syncTimerId = window.setInterval(() => {
        void this.triggerManualSync(true);
      }, this.settings.autoSyncIntervalSec * 1000);
    }

    this.updateStatusBar(this.settings.activeVault ? "Connected" : "Signed in");
  }

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    if (!this.settings.syncStore || !this.settings.syncStore.files) {
      this.settings.syncStore = { files: {} };
    }
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }
}

class EntraCredentialsModal extends Modal {
  private username = "";
  private password = "";

  constructor(
    app: App,
    private plugin: EntraSharePointSyncPlugin,
    private onSignedIn: () => void
  ) {
    super(app);
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h2", { text: "Sign in with Microsoft Entra ID" });
    contentEl.createEl("p", {
      text: "Authenticate with your Enterprise Entra ID account to activate Obsidian Enterprise Sync.",
    });

    new Setting(contentEl).setName("Work / School Account (UPN)").addText((text) =>
      text
        .setPlaceholder("user@yourdomain.onmicrosoft.com")
        .onChange((val) => {
          this.username = val;
        })
    );

    new Setting(contentEl).setName("Password").addText((text) => {
      text.inputEl.type = "password";
      text.setPlaceholder("••••••••").onChange((val) => {
        this.password = val;
      });
    });

    new Setting(contentEl).addButton((btn) =>
      btn
        .setButtonText("Sign In")
        .setCta()
        .onClick(async () => {
          try {
            const tokens = await this.plugin.auth.signInWithCredentials(
              this.username,
              this.password
            );
            this.plugin.settings.tokens = tokens;
            await this.plugin.saveSettings();
            this.plugin.reinitClients();
            new Notice(
              `Signed in as ${tokens.userDisplayName || tokens.userEmail}!`
            );
            this.onSignedIn();
            this.close();
          } catch (err) {
            new Notice(`Sign-in failed: ${String(err)}`);
          }
        })
    );
  }
}

class EntraSyncSettingTab extends PluginSettingTab {
  private availableVaults: RemoteVaultTarget[] = [];
  private newVaultName = "";

  constructor(app: App, private plugin: EntraSharePointSyncPlugin) {
    super(app, plugin);
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", {
      text: "Enterprise Sync (Entra ID & SharePoint / OneDrive)",
    });

    // Enterprise Account Status Card
    const card = containerEl.createDiv({ cls: "entra-sync-card" });
    if (this.plugin.settings.tokens) {
      const user =
        this.plugin.settings.tokens.userDisplayName ||
        this.plugin.settings.tokens.userEmail ||
        "Authenticated User";
      const email = this.plugin.settings.tokens.userEmail || "";
      card.createEl("div", {
        cls: "entra-sync-badge",
        text: "✓ Obsidian Enterprise Sync Active (Entra ID Verified)",
      });
      card.createEl("p", {
        text: `Signed in as: ${user}${email && email !== user ? ` (${email})` : ""}`,
      });
    } else {
      card.createEl("p", {
        text: "Sign in with your Microsoft Entra ID or OneDrive account to unlock multi-user shared vaults, 3-way conflict merging, and version history.",
      });
    }

    new Setting(containerEl)
      .setName("Tenant ID")
      .setDesc(
        'Use "common" for Consumer OneDrive + Multitenant dev, or your Enterprise Tenant UUID.'
      )
      .addText((text) =>
        text.setValue(this.plugin.settings.tenantId).onChange(async (val) => {
          this.plugin.settings.tenantId = val.trim() || "common";
          await this.plugin.saveSettings();
          this.plugin.reinitClients();
        })
      );

    new Setting(containerEl)
      .setName("Entra App Client ID")
      .setDesc("Application (client) ID from Microsoft Entra App Registrations.")
      .addText((text) =>
        text.setValue(this.plugin.settings.clientId).onChange(async (val) => {
          this.plugin.settings.clientId = val.trim();
          await this.plugin.saveSettings();
          this.plugin.reinitClients();
        })
      );

    new Setting(containerEl)
      .setName("OneDrive Vaults Folder Path")
      .setDesc(
        "Parent folder in OneDrive (e.g. ObsidianVaults or Documents/Documents)"
      )
      .addText((text) =>
        text
          .setValue(this.plugin.settings.oneDriveVaultsFolder)
          .onChange(async (val) => {
            this.plugin.settings.oneDriveVaultsFolder =
              val.trim() || "ObsidianVaults";
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Microsoft Graph Base URL")
      .setDesc("Default: https://graph.microsoft.com/v1.0")
      .addText((text) =>
        text.setValue(this.plugin.settings.graphBaseUrl).onChange(async (val) => {
          this.plugin.settings.graphBaseUrl =
            val.trim() || "https://graph.microsoft.com/v1.0";
          await this.plugin.saveSettings();
          this.plugin.reinitClients();
        })
      );

    const authSetting = new Setting(containerEl)
      .setName("Microsoft / Entra ID Account")
      .setDesc(
        this.plugin.settings.tokens
          ? `Connected as ${
              this.plugin.settings.tokens.userEmail ||
              this.plugin.settings.tokens.userDisplayName
            }`
          : "Not signed in"
      );

    authSetting.addButton((btn) =>
      btn.setButtonText("Sign in (Browser PKCE)").onClick(async () => {
        const url = await this.plugin.auth.startLogin();
        window.open(url);
      })
    );

    authSetting.addButton((btn) =>
      btn
        .setButtonText("Sign in (Entra Credentials)")
        .setCta()
        .onClick(() => {
          new EntraCredentialsModal(this.app, this.plugin, () => {
            this.display();
          }).open();
        })
    );

    if (this.plugin.settings.tokens) {
      authSetting.addButton((btn) =>
        btn
          .setButtonText("Sign Out")
          .setWarning()
          .onClick(async () => {
            this.plugin.settings.tokens = undefined;
            this.plugin.settings.activeVault = undefined;
            this.plugin.settings.syncStore = { files: {} };
            await this.plugin.saveSettings();
            this.plugin.reinitClients();
            this.display();
          })
      );

      containerEl.createEl("h3", { text: "Shared Vault & Collaboration" });

      new Setting(containerEl)
        .setName("Active Remote Vault")
        .setDesc(
          this.plugin.settings.activeVault
            ? `${this.plugin.settings.activeVault.name} (Drive: ${this.plugin.settings.activeVault.driveId})`
            : "No remote vault connected yet"
        )
        .addButton((btn) =>
          btn.setButtonText("Discover Vaults").onClick(async () => {
            try {
              this.availableVaults =
                await this.plugin.graph.listAvailableVaults(
                  this.plugin.settings.includeSitesScope,
                  this.plugin.settings.oneDriveVaultsFolder
                );
              new Notice(`Found ${this.availableVaults.length} vault(s).`);
              this.display();
            } catch (err) {
              new Notice(`Failed to list vaults: ${String(err)}`);
            }
          })
        )
        .addButton((btn) =>
          btn
            .setButtonText("Sync Now")
            .setCta()
            .onClick(async () => {
              await this.plugin.triggerManualSync();
            })
        );

      if (this.availableVaults.length > 0) {
        new Setting(containerEl)
          .setName("Select Discovered Vault")
          .setDesc("Choose a personal OneDrive folder, Shared with Me folder, or SharePoint library")
          .addDropdown((drop) => {
            drop.addOption("", "-- Select a vault --");
            for (const v of this.availableVaults) {
              drop.addOption(`${v.driveId}:::${v.rootItemId}`, v.name);
            }
            if (this.plugin.settings.activeVault) {
              drop.setValue(
                `${this.plugin.settings.activeVault.driveId}:::${this.plugin.settings.activeVault.rootItemId}`
              );
            }
            drop.onChange(async (val) => {
              const found = this.availableVaults.find(
                (v) => `${v.driveId}:::${v.rootItemId}` === val
              );
              if (found) {
                this.plugin.settings.activeVault = found;
                this.plugin.settings.syncStore = { files: {} };
                await this.plugin.saveSettings();
                this.plugin.reinitClients();
                await this.plugin.triggerManualSync();
                this.display();
              }
            });
          });
      }

      new Setting(containerEl)
        .setName("Create / Connect Vault Folder")
        .setDesc(
          `Create or connect a vault folder inside '${this.plugin.settings.oneDriveVaultsFolder}'`
        )
        .addText((text) =>
          text
            .setPlaceholder("e.g. ProductArchitecture")
            .setValue(this.newVaultName)
            .onChange((val) => {
              this.newVaultName = val;
            })
        )
        .addButton((btn) =>
          btn.setButtonText("Create & Connect Vault").onClick(async () => {
            if (!this.newVaultName.trim()) {
              new Notice("Enter a vault name first.");
              return;
            }
            try {
              const created = await this.plugin.graph.createOrGetVaultFolder(
                this.newVaultName.trim(),
                this.plugin.settings.oneDriveVaultsFolder
              );
              this.plugin.settings.activeVault = created;
              this.plugin.settings.syncStore = { files: {} };
              await this.plugin.saveSettings();
              this.plugin.reinitClients();
              await this.plugin.triggerManualSync();
              new Notice(`Connected to vault: ${created.name}`);
              this.display();
            } catch (err) {
              new Notice(`Failed to create vault: ${String(err)}`);
            }
          })
        );

      if (this.plugin.settings.activeVault) {
        new Setting(containerEl)
          .setName("Per-Vault Sharing (OneNote / SharePoint style)")
          .setDesc("Invite colleagues by email and manage access permissions")
          .addButton((btn) =>
            btn.setButtonText("Share Vault...").onClick(() => {
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
            btn.setButtonText("Active Note Version History").onClick(() => {
              this.plugin.openActiveFileHistory();
            })
          );
      }
    }

    // Troubleshooting & Sync Logs section (always visible)
    containerEl.createEl("h3", { text: "Troubleshooting & Sync Logs" });

    if (this.plugin.settings.tokens) {
      const expDate = new Date(this.plugin.settings.tokens.expiresAt).toLocaleString();
      const hasRefresh = Boolean(this.plugin.settings.tokens.refreshToken);
      new Setting(containerEl)
        .setName("OAuth2 Token & Delta Diagnostics")
        .setDesc(
          `Access token expires: ${expDate} • Refresh token present: ${
            hasRefresh ? "Yes" : "No"
          }`
        )
        .addButton((btn) =>
          btn.setButtonText("Force Token Refresh").onClick(async () => {
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
          btn.setButtonText("Reset Delta & Resync").onClick(async () => {
            this.plugin.settings.syncStore.deltaLink = undefined;
            await this.plugin.saveSettings();
            this.plugin.logger.info("Cleared cached deltaLink; triggering full delta sync");
            await this.plugin.triggerManualSync();
            this.display();
          })
        );
    }

    const logBoxContainer = containerEl.createDiv({ cls: "entra-sync-card" });
    const logPre = logBoxContainer.createEl("pre", {
      cls: "entra-sync-version-preview entra-sync-log-viewer",
      text: this.plugin.logger.formatText(),
    });

    new Setting(containerEl)
      .setName("Diagnostic Log Actions")
      .setDesc("Also logged to Developer Console (Ctrl+Shift+I / Cmd+Opt+I) with prefix [EntraSync]")
      .addButton((btn) =>
        btn.setButtonText("Refresh Logs").onClick(() => {
          logPre.setText(this.plugin.logger.formatText());
        })
      )
      .addButton((btn) =>
        btn.setButtonText("Copy Logs").onClick(async () => {
          await navigator.clipboard.writeText(this.plugin.logger.formatText());
          new Notice("Copied Entra Sync diagnostic logs to clipboard.");
        })
      )
      .addButton((btn) =>
        btn.setButtonText("Clear Logs").onClick(() => {
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
    contentEl.createEl("h2", { text: "Entra Sync Diagnostic Logs" });
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
          .setButtonText("Copy to Clipboard")
          .setCta()
          .onClick(async () => {
            await navigator.clipboard.writeText(this.plugin.logger.formatText());
            new Notice("Copied diagnostic logs to clipboard.");
          })
      );
  }
}
