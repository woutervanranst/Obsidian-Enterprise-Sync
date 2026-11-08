import { App, Modal, Notice, Setting } from "obsidian";
import { GraphVaultClient, RemoteVaultTarget, VaultPermission } from "./graph";

export class ShareVaultModal extends Modal {
  private recipientEmail = "";
  private role: "write" | "read" = "write";

  constructor(
    app: App,
    private target: RemoteVaultTarget,
    private graph: GraphVaultClient
  ) {
    super(app);
  }

  async onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass("entra-sync-share-modal");

    contentEl.createEl("h2", { text: `Share Vault: ${this.target.name}` });
    contentEl.createEl("p", {
      text: "Invite enterprise colleagues (Entra ID / SharePoint) or Microsoft accounts to collaborate in this vault.",
    });

    new Setting(contentEl)
      .setName("Colleague Email(s)")
      .setDesc("Comma-separated email addresses (e.g. bob@woutervanranstgmail.onmicrosoft.com)")
      .addText((text) =>
        text
          .setPlaceholder("colleague@company.com")
          .setValue(this.recipientEmail)
          .onChange((val) => {
            this.recipientEmail = val;
          })
      );

    new Setting(contentEl)
      .setName("Permission Role")
      .setDesc("Write allows editing & creating notes; Read allows viewing only.")
      .addDropdown((drop) =>
        drop
          .addOption("write", "Can edit (Write)")
          .addOption("read", "Can view (Read)")
          .setValue(this.role)
          .onChange((val) => {
            this.role = val as "write" | "read";
          })
      );

    const permissionsContainer = contentEl.createDiv({ cls: "entra-sync-card" });
    permissionsContainer.createEl("h4", { text: "Current Vault Collaborators" });
    const permsListEl = permissionsContainer.createDiv({ cls: "entra-sync-collaborators-list" });

    const refreshPermissions = async () => {
      permsListEl.empty();
      const perms: VaultPermission[] = await this.graph.listPermissions(this.target);
      if (perms.length === 0) {
        permsListEl.createEl("p", { text: "Only you currently have access." });
        return;
      }
      for (const p of perms) {
        const name =
          p.grantedToV2?.user?.displayName ||
          p.grantedToV2?.user?.email ||
          p.invitation?.email ||
          "Collaborator";
        const email =
          p.grantedToV2?.user?.email || p.invitation?.email || "";
        const roles = (p.roles || []).join(", ");
        permsListEl.createEl("div", {
          cls: "entra-sync-collaborator-row",
          text: `• ${name}${email && email !== name ? ` (${email})` : ""} — [${roles}]`,
        });
      }
    };

    new Setting(contentEl).addButton((btn) =>
      btn
        .setButtonText("Invite to Vault")
        .setCta()
        .onClick(async () => {
          const emails = this.recipientEmail
            .split(",")
            .map((e) => e.trim())
            .filter(Boolean);
          if (emails.length === 0) {
            new Notice("Please enter at least one email address.");
            return;
          }
          try {
            await this.graph.inviteCollaborators(this.target, emails, this.role);
            new Notice(`Shared vault '${this.target.name}' with ${emails.join(", ")}`);
            this.recipientEmail = "";
            await refreshPermissions();
          } catch (err) {
            new Notice(`Failed to share vault: ${String(err)}`);
          }
        })
    );

    await refreshPermissions();
  }
}
