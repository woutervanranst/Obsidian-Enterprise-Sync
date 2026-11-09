import { App, Modal, Notice, Setting, TFile } from "obsidian";
import { DriveItemVersion, GraphVaultClient } from "./graph";

export class VersionHistoryModal extends Modal {
  constructor(
    app: App,
    private file: TFile,
    private driveId: string,
    private itemId: string,
    private graph: GraphVaultClient,
    private onRestored?: (newContent: string) => Promise<void>
  ) {
    super(app);
  }

  async onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass("entra-sync-history-modal");
    contentEl.createEl("h2", { text: `Version history: ${this.file.name}` });
    const loading = contentEl.createEl("p", {
      text: "Loading version history from SharePoint / OneDrive...",
    });

    try {
      const versions: DriveItemVersion[] = await this.graph.listVersions(
        this.driveId,
        this.itemId
      );
      loading.remove();

      if (versions.length === 0) {
        contentEl.createEl("p", { text: "No versions found for this note." });
        return;
      }

      const previewContainer = contentEl.createDiv({ cls: "entra-sync-card" });
      previewContainer.createEl("strong", { text: "Version preview" });
      const previewEl = previewContainer.createEl("pre", {
        cls: "entra-sync-version-preview",
        text: "Select preview on any version below to inspect its content.",
      });

      const listContainer = contentEl.createDiv({ cls: "entra-sync-version-list" });

      for (const v of versions) {
        const author =
          v.lastModifiedBy?.user?.displayName ||
          v.lastModifiedBy?.user?.email ||
          "Unknown author";
        const email = v.lastModifiedBy?.user?.email
          ? ` (${v.lastModifiedBy.user.email})`
          : "";
        const dateStr = new Date(v.lastModifiedDateTime).toLocaleString();
        const sizeStr = typeof v.size === "number" ? ` • ${v.size} bytes` : "";

        new Setting(listContainer)
          .setName(`Version ${v.id} — ${author}${email}`)
          .setDesc(`${dateStr}${sizeStr}`)
          .addButton((btn) =>
            btn.setButtonText("Preview").onClick(async () => {
              previewEl.setText(`Loading version ${v.id}...`);
              const text = await this.graph.downloadVersionText(
                this.driveId,
                this.itemId,
                v.id
              );
              previewEl.setText(text);
            })
          )
          .addButton((btn) =>
            btn
              .setButtonText("Restore")
              .setCta()
              .onClick(async () => {
                const text = await this.graph.downloadVersionText(
                  this.driveId,
                  this.itemId,
                  v.id
                );
                await this.app.vault.modify(this.file, text);
                if (this.onRestored) {
                  await this.onRestored(text);
                }
                new Notice(`Restored version ${v.id} by ${author}`);
                this.close();
              })
          );
      }
    } catch (err) {
      loading.setText(`Failed to load versions: ${String(err)}`);
    }
  }
}
