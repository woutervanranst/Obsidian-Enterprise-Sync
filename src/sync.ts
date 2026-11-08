import { TFile, TFolder, Vault } from "obsidian";
import { diff3Merge } from "node-diff3";
import { DeltaDriveItem, GraphVaultClient, RemoteVaultTarget } from "./graph";

export interface SyncFileState {
  itemId: string;
  eTag: string;
  /** Last synced base text used for 3-way merge when concurrent edits occur */
  baseText: string;
  lastSyncedMtime: number;
  lastModifiedBy?: string;
}

export interface SyncStore {
  deltaLink?: string;
  files: Record<string, SyncFileState>;
}

export interface SyncSummary {
  pulled: number;
  pushed: number;
  merged: number;
  conflicts: number;
}

export class CollaborativeSyncEngine {
  private isSyncing = false;
  private suppressModifyForPaths = new Set<string>();

  constructor(
    private vault: Vault,
    private graph: GraphVaultClient,
    private target: RemoteVaultTarget,
    private store: SyncStore,
    private saveStore: () => Promise<void>,
    private onStatusChange?: (statusText: string) => void
  ) {}

  isPathSuppressed(path: string): boolean {
    return this.suppressModifyForPaths.has(path);
  }

  /**
   * Full bidirectional sync: pull delta changes from SharePoint/OneDrive, then push local changes.
   */
  async syncAll(): Promise<SyncSummary> {
    if (this.isSyncing) {
      return { pulled: 0, pushed: 0, merged: 0, conflicts: 0 };
    }
    this.isSyncing = true;
    this.onStatusChange?.("Syncing...");

    const summary: SyncSummary = { pulled: 0, pushed: 0, merged: 0, conflicts: 0 };
    try {
      const pullStats = await this.pullRemoteChanges();
      summary.pulled += pullStats.pulled;
      summary.merged += pullStats.merged;
      summary.conflicts += pullStats.conflicts;

      const mdFiles = this.vault.getMarkdownFiles();
      for (const file of mdFiles) {
        const pushStats = await this.pushMarkdownFile(file);
        summary.pushed += pushStats.pushed;
        summary.merged += pushStats.merged;
        summary.conflicts += pushStats.conflicts;
      }

      this.onStatusChange?.(`Synced (${new Date().toLocaleTimeString()})`);
      return summary;
    } catch (err) {
      this.onStatusChange?.("Sync Error");
      throw err;
    } finally {
      this.isSyncing = false;
    }
  }

  /**
   * Pull incremental changes from SharePoint / OneDrive via Microsoft Graph /delta.
   */
  async pullRemoteChanges(): Promise<{ pulled: number; merged: number; conflicts: number }> {
    let pulled = 0;
    let merged = 0;
    let conflicts = 0;

    const { items, nextDeltaLink } = await this.graph.fetchDelta(
      this.target,
      this.store.deltaLink
    );

    for (const item of items) {
      if (item.folder || item.id === this.target.rootItemId) {
        continue;
      }

      const relPath = this.resolveRelativePath(item);
      if (!relPath || !relPath.endsWith(".md")) {
        continue;
      }

      const existingState = this.store.files[relPath];

      if (item.deleted) {
        const existingFile = this.vault.getAbstractFileByPath(relPath);
        if (existingFile instanceof TFile) {
          await this.vault.delete(existingFile);
        }
        delete this.store.files[relPath];
        pulled++;
        continue;
      }

      // Skip if eTag matches what we already have synced
      if (existingState && item.eTag && existingState.eTag === item.eTag) {
        continue;
      }

      const remote = await this.graph.downloadFileText(this.target.driveId, item.id);
      const localAbstract = this.vault.getAbstractFileByPath(relPath);

      if (localAbstract instanceof TFile) {
        const localText = await this.vault.read(localAbstract);

        if (!existingState || localText === existingState.baseText) {
          // Clean fast-forward update from colleague
          if (localText !== remote.text) {
            await this.writeLocalSuppressed(localAbstract, remote.text);
          }
          this.store.files[relPath] = {
            itemId: item.id,
            eTag: remote.eTag || item.eTag || "",
            baseText: remote.text,
            lastSyncedMtime: localAbstract.stat.mtime,
            lastModifiedBy: remote.lastModifiedBy,
          };
          pulled++;
        } else if (localText !== remote.text) {
          // Both local and remote changed since baseText -> 3-way merge
          const base = existingState.baseText;
          const { text: mergedText, hadConflict } = this.threeWayMerge(
            base,
            localText,
            remote.text
          );
          await this.writeLocalSuppressed(localAbstract, mergedText);

          const uploadRes = await this.graph.uploadFileText(
            this.target,
            relPath,
            mergedText,
            remote.eTag
          );
          const finalItem = uploadRes.item ?? item;
          this.store.files[relPath] = {
            itemId: finalItem.id,
            eTag: finalItem.eTag ?? remote.eTag,
            baseText: mergedText,
            lastSyncedMtime: localAbstract.stat.mtime,
            lastModifiedBy: remote.lastModifiedBy,
          };
          merged++;
          if (hadConflict) conflicts++;
        }
      } else {
        // New file from SharePoint / OneDrive
        await this.ensureParentFolders(relPath);
        this.suppressModifyForPaths.add(relPath);
        try {
          const created = await this.vault.create(relPath, remote.text);
          this.store.files[relPath] = {
            itemId: item.id,
            eTag: remote.eTag || item.eTag || "",
            baseText: remote.text,
            lastSyncedMtime: created.stat.mtime,
            lastModifiedBy: remote.lastModifiedBy,
          };
          pulled++;
        } finally {
          this.suppressModifyForPaths.delete(relPath);
        }
      }
    }

    if (nextDeltaLink) {
      this.store.deltaLink = nextDeltaLink;
    }
    await this.saveStore();
    return { pulled, merged, conflicts };
  }

  /**
   * Push a locally modified markdown file to SharePoint/OneDrive.
   * If a colleague modified it since our last sync (HTTP 412), perform an automatic
   * 3-way merge using our stored baseText and re-upload cleanly.
   */
  async pushMarkdownFile(
    file: TFile
  ): Promise<{ pushed: number; merged: number; conflicts: number }> {
    if (this.suppressModifyForPaths.has(file.path)) {
      return { pushed: 0, merged: 0, conflicts: 0 };
    }

    const localText = await this.vault.read(file);
    const state = this.store.files[file.path];

    if (state && localText === state.baseText) {
      return { pushed: 0, merged: 0, conflicts: 0 };
    }

    const res = await this.graph.uploadFileText(
      this.target,
      file.path,
      localText,
      state?.eTag
    );

    if (res.status === 412 && state?.itemId) {
      // Concurrent edit detected! Fetch remote version and run 3-way diff3 merge
      const remote = await this.graph.downloadFileText(this.target.driveId, state.itemId);
      const { text: mergedText, hadConflict } = this.threeWayMerge(
        state.baseText,
        localText,
        remote.text
      );

      await this.writeLocalSuppressed(file, mergedText);
      const retry = await this.graph.uploadFileText(
        this.target,
        file.path,
        mergedText,
        remote.eTag
      );

      if (retry.item) {
        this.store.files[file.path] = {
          itemId: retry.item.id,
          eTag: retry.item.eTag ?? "",
          baseText: mergedText,
          lastSyncedMtime: file.stat.mtime,
          lastModifiedBy: remote.lastModifiedBy,
        };
        await this.saveStore();
      }
      return { pushed: 1, merged: 1, conflicts: hadConflict ? 1 : 0 };
    }

    if (res.item) {
      this.store.files[file.path] = {
        itemId: res.item.id,
        eTag: res.item.eTag ?? "",
        baseText: localText,
        lastSyncedMtime: file.stat.mtime,
      };
      await this.saveStore();
      return { pushed: 1, merged: 0, conflicts: 0 };
    }

    return { pushed: 0, merged: 0, conflicts: 0 };
  }

  private async writeLocalSuppressed(file: TFile, content: string): Promise<void> {
    this.suppressModifyForPaths.add(file.path);
    try {
      await this.vault.modify(file, content);
    } finally {
      this.suppressModifyForPaths.delete(file.path);
    }
  }

  private async ensureParentFolders(relPath: string): Promise<void> {
    const parts = relPath.split("/");
    parts.pop();
    let current = "";
    for (const part of parts) {
      current = current ? `${current}/${part}` : part;
      const existing = this.vault.getAbstractFileByPath(current);
      if (!existing) {
        await this.vault.createFolder(current);
      } else if (!(existing instanceof TFolder)) {
        break;
      }
    }
  }

  private resolveRelativePath(item: DeltaDriveItem): string {
    const parentPath = item.parentReference?.path;
    if (!parentPath) {
      return item.name;
    }
    // parentReference.path looks like "/drive/root:/ObsidianVaults/SharedVault/Subfolder"
    // or "/drives/{id}/root:/SharedVault"
    const colonIdx = parentPath.indexOf(":");
    if (colonIdx === -1) {
      return item.name;
    }
    const afterColon = parentPath.slice(colonIdx + 1).replace(/^\/+/, "");
    if (!afterColon) {
      return item.name;
    }
    // Strip the vault folder prefix if present
    const vaultFolderName = this.target.name.replace(/\s*\(.*?\)\s*$/, "").trim();
    const segments = afterColon.split("/").filter(Boolean);
    const vaultIdx = segments.lastIndexOf(vaultFolderName);
    const relSegments = vaultIdx !== -1 ? segments.slice(vaultIdx + 1) : [];
    return relSegments.length > 0 ? `${relSegments.join("/")}/${item.name}` : item.name;
  }

  /**
   * Line-based 3-way merge. Non-overlapping edits merge silently;
   * overlapping same-line edits insert standard conflict markers without duplicating files.
   */
  threeWayMerge(
    base: string,
    local: string,
    remote: string
  ): { text: string; hadConflict: boolean } {
    const result = diff3Merge(
      local.split("\n"),
      base.split("\n"),
      remote.split("\n")
    );

    const out: string[] = [];
    let hadConflict = false;
    for (const chunk of result) {
      if (chunk.ok) {
        out.push(...chunk.ok);
      } else if (chunk.conflict) {
        hadConflict = true;
        out.push(
          "<<<<<<< Local",
          ...chunk.conflict.a,
          "=======",
          ...chunk.conflict.b,
          ">>>>>>> Remote (SharePoint)"
        );
      }
    }
    return { text: out.join("\n"), hadConflict };
  }
}
