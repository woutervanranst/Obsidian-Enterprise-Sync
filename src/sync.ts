import { TFile, TFolder, Vault } from "obsidian";
import { diff3Merge } from "node-diff3";
import { DeltaDriveItem, GraphVaultClient, RemoteVaultTarget } from "./graph";
import { SyncLogger, defaultLogger } from "./logger";

export interface SyncFileState {
  itemId: string;
  eTag: string;
  /** Last synced base text used for 3-way merge when concurrent edits occur (empty for binary files) */
  baseText: string;
  isBinary?: boolean;
  /** Fast hex digest + byte length for binary attachments */
  binaryHash?: string;
  lastSyncedMtime: number;
  lastModifiedBy?: string;
}

export interface SyncStore {
  deltaLink?: string;
  files: Record<string, SyncFileState>;
  /** Tracks remote folder relative paths -> itemId for folder deletion handling */
  folders?: Record<string, string>;
}

export interface SyncSummary {
  pulled: number;
  pushed: number;
  deleted?: number;
  merged: number;
  conflicts: number;
}

const TEXT_EXTENSIONS = new Set([
  "md",
  "canvas",
  "txt",
  "json",
  "csv",
  "yaml",
  "yml",
  "xml",
  "svg",
  "html",
  "css",
  "js",
  "ts",
]);

export function isTextPath(relPath: string): boolean {
  const name = relPath.split("/").pop() || "";
  const dotIdx = name.lastIndexOf(".");
  if (dotIdx === -1) return true;
  const ext = name.slice(dotIdx + 1).toLowerCase();
  return TEXT_EXTENSIONS.has(ext);
}

export function isIgnoredVaultPath(relPath: string): boolean {
  const normalized = relPath.replace(/^\/+/, "");
  if (!normalized) return true;
  return (
    normalized === ".obsidian" ||
    normalized.startsWith(".obsidian/") ||
    normalized.endsWith("/.DS_Store") ||
    normalized === ".DS_Store"
  );
}

export function hashBinaryBuffer(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  for (let i = 0; i < bytes.length; i++) {
    h1 ^= bytes[i];
    h1 = Math.imul(h1, 0x01000193);
    h2 ^= bytes[i] + (i & 0xff);
    h2 = Math.imul(h2, 0x85ebca6b);
  }
  return `${bytes.byteLength}:${(h1 >>> 0).toString(16).padStart(8, "0")}${(h2 >>> 0)
    .toString(16)
    .padStart(8, "0")}`;
}

export class CollaborativeSyncEngine {
  private isSyncing = false;
  private suppressModifyForPaths = new Set<string>();
  private suppressDeleteForPaths = new Set<string>();

  constructor(
    private vault: Vault,
    private graph: GraphVaultClient,
    private target: RemoteVaultTarget,
    private store: SyncStore,
    private saveStore: () => Promise<void>,
    private onStatusChange?: (statusText: string) => void,
    private logger: SyncLogger = defaultLogger
  ) {
    if (!this.store.folders) {
      this.store.folders = {};
    }
  }

  isPathSuppressed(path: string): boolean {
    return this.suppressModifyForPaths.has(path);
  }

  isDeleteSuppressed(path: string): boolean {
    return this.suppressDeleteForPaths.has(path);
  }

  private getSyncableVaultFiles(): TFile[] {
    const allFiles: TFile[] =
      typeof (this.vault as unknown as { getFiles?: () => TFile[] }).getFiles ===
      "function"
        ? (this.vault as unknown as { getFiles: () => TFile[] }).getFiles()
        : this.vault.getMarkdownFiles();
    return allFiles.filter((f) => !isIgnoredVaultPath(f.path));
  }

  /**
   * Full bidirectional sync:
   * 1. Pull delta changes (new/updated/deleted files & folders, text + binary attachments) from SharePoint/OneDrive.
   * 2. Reconcile local deletions (files & folders tracked in store that were deleted locally) -> delete on SharePoint/OneDrive.
   * 3. Push local new/modified files (both Markdown/text notes and binary attachments).
   */
  async syncAll(): Promise<SyncSummary> {
    if (this.isSyncing) {
      return { pulled: 0, pushed: 0, deleted: 0, merged: 0, conflicts: 0 };
    }
    this.isSyncing = true;
    this.onStatusChange?.("Syncing...");
    if (!this.store.folders) {
      this.store.folders = {};
    }
    this.logger.info(`Starting syncAll for vault '${this.target.name}'`, {
      driveId: this.target.driveId,
      rootItemId: this.target.rootItemId,
      hasDeltaLink: Boolean(this.store.deltaLink),
    });

    const summary: SyncSummary = {
      pulled: 0,
      pushed: 0,
      deleted: 0,
      merged: 0,
      conflicts: 0,
    };
    try {
      const pullStats = await this.pullRemoteChanges();
      summary.pulled += pullStats.pulled;
      summary.merged += pullStats.merged;
      summary.conflicts += pullStats.conflicts;

      // Reconcile local file deletions (tracked files that no longer exist locally)
      for (const [trackedPath, state] of Object.entries({ ...this.store.files })) {
        const localExists = this.vault.getAbstractFileByPath(trackedPath);
        if (!localExists) {
          this.logger.info(`Local file deleted: '${trackedPath}' -> deleting remote item ${state.itemId}`);
          if (state.itemId) {
            await this.graph.deleteItem(this.target.driveId, state.itemId);
          }
          delete this.store.files[trackedPath];
          summary.deleted = (summary.deleted || 0) + 1;
        }
      }

      // Reconcile local folder deletions (tracked folders that no longer exist locally)
      for (const [trackedFolder, folderItemId] of Object.entries({
        ...(this.store.folders || {}),
      })) {
        const localFolder = this.vault.getAbstractFileByPath(trackedFolder);
        if (!localFolder) {
          this.logger.info(
            `Local folder deleted: '${trackedFolder}' -> deleting remote folder ${folderItemId}`
          );
          if (folderItemId) {
            await this.graph.deleteItem(this.target.driveId, folderItemId);
          }
          delete this.store.folders![trackedFolder];
        }
      }
      if ((summary.deleted || 0) > 0) {
        await this.saveStore();
      }

      const files = this.getSyncableVaultFiles();
      for (const file of files) {
        const pushStats = await this.pushFile(file);
        summary.pushed += pushStats.pushed;
        summary.merged += pushStats.merged;
        summary.conflicts += pushStats.conflicts;
      }

      this.logger.info("Completed syncAll successfully", summary);
      this.onStatusChange?.(`Synced (${new Date().toLocaleTimeString()})`);
      return summary;
    } catch (err) {
      this.logger.error("syncAll failed", String(err));
      this.onStatusChange?.("Sync Error");
      throw err;
    } finally {
      this.isSyncing = false;
    }
  }

  /**
   * Immediately propagate a local file or folder deletion to SharePoint / OneDrive.
   */
  async deleteLocalPath(relPath: string, isFolder = false): Promise<number> {
    if (this.suppressDeleteForPaths.has(relPath) || isIgnoredVaultPath(relPath)) {
      return 0;
    }
    if (!this.store.folders) {
      this.store.folders = {};
    }

    let deletedCount = 0;

    if (isFolder) {
      const folderPrefix = `${relPath}/`;
      const folderItemId = this.store.folders[relPath];
      for (const [trackedPath, state] of Object.entries({ ...this.store.files })) {
        if (trackedPath === relPath || trackedPath.startsWith(folderPrefix)) {
          if (state.itemId) {
            await this.graph.deleteItem(this.target.driveId, state.itemId);
          }
          delete this.store.files[trackedPath];
          deletedCount++;
        }
      }
      for (const [trackedFolder, id] of Object.entries({ ...this.store.folders })) {
        if (trackedFolder === relPath || trackedFolder.startsWith(folderPrefix)) {
          if (id && trackedFolder !== relPath) {
            await this.graph.deleteItem(this.target.driveId, id);
          }
          delete this.store.folders[trackedFolder];
        }
      }
      if (folderItemId) {
        await this.graph.deleteItem(this.target.driveId, folderItemId);
      } else {
        await this.graph.deleteItemByPath(this.target, relPath);
      }
      deletedCount++;
    } else {
      const state = this.store.files[relPath];
      if (state?.itemId) {
        await this.graph.deleteItem(this.target.driveId, state.itemId);
        delete this.store.files[relPath];
        deletedCount++;
      } else {
        await this.graph.deleteItemByPath(this.target, relPath);
      }
    }

    if (deletedCount > 0) {
      await this.saveStore();
    }
    return deletedCount;
  }

  /**
   * Handle a local file or folder rename/move in Obsidian.
   */
  async renameLocalPath(
    oldPath: string,
    newAbstract: TFile | TFolder
  ): Promise<void> {
    if (isIgnoredVaultPath(newAbstract.path)) return;
    const isFolder = newAbstract instanceof TFolder;
    await this.deleteLocalPath(oldPath, isFolder);

    if (newAbstract instanceof TFile) {
      await this.pushFile(newAbstract);
    } else if (newAbstract instanceof TFolder) {
      const folderPrefix = `${newAbstract.path}/`;
      for (const file of this.getSyncableVaultFiles()) {
        if (file.path.startsWith(folderPrefix)) {
          await this.pushFile(file);
        }
      }
    }
  }

  private findTrackedFilePathByItemId(itemId: string): string | undefined {
    for (const [path, state] of Object.entries(this.store.files)) {
      if (state.itemId === itemId) return path;
    }
    return undefined;
  }

  private findTrackedFolderPathByItemId(itemId: string): string | undefined {
    for (const [path, id] of Object.entries(this.store.folders || {})) {
      if (id === itemId) return path;
    }
    return undefined;
  }

  /**
   * Pull incremental changes from SharePoint / OneDrive via Microsoft Graph /delta.
   */
  async pullRemoteChanges(): Promise<{ pulled: number; merged: number; conflicts: number }> {
    let pulled = 0;
    let merged = 0;
    let conflicts = 0;
    if (!this.store.folders) {
      this.store.folders = {};
    }

    const { items, nextDeltaLink } = await this.graph.fetchDelta(
      this.target,
      this.store.deltaLink
    );

    for (const item of items) {
      if (item.id === this.target.rootItemId) {
        continue;
      }

      // Handle remote deletions (both files and folders, including tombstones where parentReference.path may be omitted)
      if (item.deleted) {
        const trackedFolderPath =
          this.findTrackedFolderPathByItemId(item.id) ||
          (item.folder ? this.resolveRelativePath(item) : undefined);

        if (trackedFolderPath && !isIgnoredVaultPath(trackedFolderPath)) {
          const prefix = `${trackedFolderPath}/`;
          for (const trackedFile of Object.keys({ ...this.store.files })) {
            if (trackedFile.startsWith(prefix)) {
              const existingChild = this.vault.getAbstractFileByPath(trackedFile);
              if (existingChild) {
                await this.deleteLocalSuppressed(existingChild);
              }
              delete this.store.files[trackedFile];
            }
          }
          for (const subFolder of Object.keys({ ...this.store.folders })) {
            if (subFolder === trackedFolderPath || subFolder.startsWith(prefix)) {
              delete this.store.folders[subFolder];
            }
          }
          const existingFolder = this.vault.getAbstractFileByPath(trackedFolderPath);
          if (existingFolder) {
            await this.deleteLocalSuppressed(existingFolder);
          }
          pulled++;
          continue;
        }

        const relPath =
          this.findTrackedFilePathByItemId(item.id) || this.resolveRelativePath(item);
        if (relPath && !isIgnoredVaultPath(relPath)) {
          const existingFile = this.vault.getAbstractFileByPath(relPath);
          if (existingFile) {
            await this.deleteLocalSuppressed(existingFile);
          }
          delete this.store.files[relPath];
          pulled++;
        }
        continue;
      }

      const relPath = this.resolveRelativePath(item);
      if (!relPath || isIgnoredVaultPath(relPath)) {
        continue;
      }

      if (item.folder) {
        this.store.folders[relPath] = item.id;
        await this.ensureFolderExists(relPath);
        continue;
      }

      const existingState = this.store.files[relPath];

      // Skip if eTag matches what we already have synced
      if (existingState && item.eTag && existingState.eTag === item.eTag) {
        continue;
      }

      if (!isTextPath(relPath)) {
        // Binary attachment (.png, .jpg, .pdf, etc.)
        const remoteBin = await this.graph.downloadFileBinary(
          this.target.driveId,
          item.id
        );
        const remoteHash = hashBinaryBuffer(remoteBin.data);
        const localAbstract = this.vault.getAbstractFileByPath(relPath);

        if (localAbstract instanceof TFile) {
          const localBuf = await this.readBinaryCompat(localAbstract);
          const localHash = hashBinaryBuffer(localBuf);

          if (
            existingState?.binaryHash &&
            localHash !== existingState.binaryHash &&
            localHash !== remoteHash
          ) {
            // Concurrent binary edit: preserve local binary attachment as a conflict copy before writing remote
            const conflictPath = this.buildBinaryConflictPath(relPath);
            await this.createBinarySuppressed(conflictPath, localBuf);
            conflicts++;
          }

          if (localHash !== remoteHash) {
            await this.writeBinarySuppressed(localAbstract, remoteBin.data);
          }
          this.store.files[relPath] = {
            itemId: item.id,
            eTag: remoteBin.eTag || item.eTag || "",
            baseText: "",
            isBinary: true,
            binaryHash: remoteHash,
            lastSyncedMtime: localAbstract.stat.mtime,
            lastModifiedBy: remoteBin.lastModifiedBy,
          };
          pulled++;
        } else {
          await this.ensureParentFolders(relPath);
          const created = await this.createBinarySuppressed(relPath, remoteBin.data);
          this.store.files[relPath] = {
            itemId: item.id,
            eTag: remoteBin.eTag || item.eTag || "",
            baseText: "",
            isBinary: true,
            binaryHash: remoteHash,
            lastSyncedMtime: created.stat.mtime,
            lastModifiedBy: remoteBin.lastModifiedBy,
          };
          pulled++;
        }
        continue;
      }

      // Text / Markdown file with 3-way diff3 merge
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
   * Push a locally modified file (Markdown/text or binary attachment) to SharePoint/OneDrive.
   */
  async pushFile(
    file: TFile
  ): Promise<{ pushed: number; merged: number; conflicts: number }> {
    if (this.suppressModifyForPaths.has(file.path) || isIgnoredVaultPath(file.path)) {
      return { pushed: 0, merged: 0, conflicts: 0 };
    }

    if (!isTextPath(file.path)) {
      return this.pushBinaryFile(file);
    }

    const localText = await this.vault.read(file);
    const state = this.store.files[file.path];

    if (state && !state.isBinary && localText === state.baseText) {
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

  /**
   * Backwards-compatible alias for pushFile().
   */
  async pushMarkdownFile(
    file: TFile
  ): Promise<{ pushed: number; merged: number; conflicts: number }> {
    return this.pushFile(file);
  }

  private async pushBinaryFile(
    file: TFile
  ): Promise<{ pushed: number; merged: number; conflicts: number }> {
    const localBuf = await this.readBinaryCompat(file);
    const localHash = hashBinaryBuffer(localBuf);
    const state = this.store.files[file.path];

    if (state?.isBinary && state.binaryHash === localHash) {
      return { pushed: 0, merged: 0, conflicts: 0 };
    }

    const res = await this.graph.uploadFileBinary(
      this.target,
      file.path,
      localBuf,
      state?.eTag
    );

    if (res.status === 412 && state?.itemId) {
      // Concurrent binary update: save local copy as conflict file and update main file to remote
      const remoteBin = await this.graph.downloadFileBinary(
        this.target.driveId,
        state.itemId
      );
      const conflictPath = this.buildBinaryConflictPath(file.path);
      const conflictFile = await this.createBinarySuppressed(conflictPath, localBuf);
      await this.writeBinarySuppressed(file, remoteBin.data);
      const remoteHash = hashBinaryBuffer(remoteBin.data);
      this.store.files[file.path] = {
        itemId: state.itemId,
        eTag: remoteBin.eTag,
        baseText: "",
        isBinary: true,
        binaryHash: remoteHash,
        lastSyncedMtime: file.stat.mtime,
        lastModifiedBy: remoteBin.lastModifiedBy,
      };
      await this.pushBinaryFile(conflictFile);
      await this.saveStore();
      return { pushed: 1, merged: 0, conflicts: 1 };
    }

    if (res.item) {
      this.store.files[file.path] = {
        itemId: res.item.id,
        eTag: res.item.eTag ?? "",
        baseText: "",
        isBinary: true,
        binaryHash: localHash,
        lastSyncedMtime: file.stat.mtime,
      };
      await this.saveStore();
      return { pushed: 1, merged: 0, conflicts: 0 };
    }

    return { pushed: 0, merged: 0, conflicts: 0 };
  }

  private buildBinaryConflictPath(relPath: string): string {
    const dotIdx = relPath.lastIndexOf(".");
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    if (dotIdx === -1) {
      return `${relPath} (Conflict ${stamp})`;
    }
    return `${relPath.slice(0, dotIdx)} (Conflict ${stamp})${relPath.slice(dotIdx)}`;
  }

  private async readBinaryCompat(file: TFile): Promise<ArrayBuffer> {
    const v = this.vault as unknown as {
      readBinary?: (f: TFile) => Promise<ArrayBuffer>;
    };
    if (typeof v.readBinary === "function") {
      return v.readBinary(file);
    }
    const text = await this.vault.read(file);
    return new TextEncoder().encode(text).buffer;
  }

  private async writeBinarySuppressed(
    file: TFile,
    data: ArrayBuffer
  ): Promise<void> {
    this.suppressModifyForPaths.add(file.path);
    try {
      const v = this.vault as unknown as {
        modifyBinary?: (f: TFile, d: ArrayBuffer) => Promise<void>;
      };
      if (typeof v.modifyBinary === "function") {
        await v.modifyBinary(file, data);
      } else {
        await this.vault.modify(file, new TextDecoder().decode(data));
      }
    } finally {
      this.suppressModifyForPaths.delete(file.path);
    }
  }

  private async createBinarySuppressed(
    relPath: string,
    data: ArrayBuffer
  ): Promise<TFile> {
    this.suppressModifyForPaths.add(relPath);
    try {
      const v = this.vault as unknown as {
        createBinary?: (p: string, d: ArrayBuffer) => Promise<TFile>;
      };
      if (typeof v.createBinary === "function") {
        return await v.createBinary(relPath, data);
      }
      return await this.vault.create(relPath, new TextDecoder().decode(data));
    } finally {
      this.suppressModifyForPaths.delete(relPath);
    }
  }

  private async deleteLocalSuppressed(abstractFile: TFile | TFolder | { path: string }): Promise<void> {
    this.suppressDeleteForPaths.add(abstractFile.path);
    try {
      await (this.vault as unknown as { delete: (f: unknown, force?: boolean) => Promise<void> }).delete(
        abstractFile,
        true
      );
    } finally {
      this.suppressDeleteForPaths.delete(abstractFile.path);
    }
  }

  private async writeLocalSuppressed(file: TFile, content: string): Promise<void> {
    this.suppressModifyForPaths.add(file.path);
    try {
      await this.vault.modify(file, content);
    } finally {
      this.suppressModifyForPaths.delete(file.path);
    }
  }

  private async ensureFolderExists(folderPath: string): Promise<void> {
    const parts = folderPath.split("/").filter(Boolean);
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

  private async ensureParentFolders(relPath: string): Promise<void> {
    const parts = relPath.split("/");
    parts.pop();
    if (parts.length > 0) {
      await this.ensureFolderExists(parts.join("/"));
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
