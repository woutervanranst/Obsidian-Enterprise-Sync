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

export function isIgnoredVaultPath(
  relPath: string,
  syncObsidianConfig = false,
  configDir = ".obsidian"
): boolean {
  const normalized = relPath.replace(/^\/+/, "");
  if (!normalized) return true;
  if (
    normalized === ".DS_Store" ||
    normalized.endsWith("/.DS_Store") ||
    normalized === ".git" ||
    normalized.startsWith(".git/") ||
    normalized === ".trash" ||
    normalized.startsWith(".trash/")
  ) {
    return true;
  }

  const cleanConfigDir = (configDir || ".obsidian").replace(/^\/+|\/+$/g, "");
  if (
    normalized === cleanConfigDir ||
    normalized.startsWith(`${cleanConfigDir}/`)
  ) {
    if (!syncObsidianConfig) {
      return true;
    }
    const subPath = normalized.slice(cleanConfigDir.length).replace(/^\/+/, "");
    // Never sync personal Entra ID tokens/sync state or per-user workspace layout
    if (
      subPath === "workspace.json" ||
      subPath === "workspace-mobile.json" ||
      subPath === "plugins/entra-sharepoint-sync/data.json" ||
      subPath.startsWith("plugins/entra-sharepoint-sync/data.json") ||
      subPath === "plugins/obsidian-entra-sync/data.json" ||
      subPath.startsWith("plugins/obsidian-entra-sync/data.json")
    ) {
      return true;
    }
    return false;
  }

  return false;
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
    private logger: SyncLogger = defaultLogger,
    private shouldSyncObsidianConfig: () => boolean = () => true
  ) {
    if (!this.store.folders) {
      this.store.folders = {};
    }
  }

  private getConfigDir(): string {
    const v = this.vault as unknown as { configDir?: string };
    return (v.configDir || ".obsidian").replace(/^\/+|\/+$/g, "");
  }

  private isConfigDirPath(relPath: string): boolean {
    const normalized = relPath.replace(/^\/+/, "");
    const cfg = this.getConfigDir();
    return normalized === cfg || normalized.startsWith(`${cfg}/`);
  }

  private isIgnored(relPath: string): boolean {
    return isIgnoredVaultPath(
      relPath,
      this.shouldSyncObsidianConfig(),
      this.getConfigDir()
    );
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
    return allFiles.filter(
      (f) => !this.isConfigDirPath(f.path) && !this.isIgnored(f.path)
    );
  }

  /**
   * Lists all syncable files inside `.obsidian/` via `vault.adapter.list()`.
   */
  private async listSyncableConfigFiles(): Promise<string[]> {
    if (!this.shouldSyncObsidianConfig()) {
      return [];
    }
    const cfgDir = this.getConfigDir();
    const adapter = (this.vault as unknown as {
      adapter?: {
        exists?: (p: string) => Promise<boolean>;
        list?: (p: string) => Promise<{ files: string[]; folders: string[] }>;
      };
    }).adapter;

    if (!adapter || typeof adapter.list !== "function") {
      return [];
    }

    const result: string[] = [];
    const queue: string[] = [cfgDir];
    const visited = new Set<string>();

    while (queue.length > 0) {
      const currentDir = queue.shift()!;
      if (visited.has(currentDir)) continue;
      visited.add(currentDir);

      try {
        if (typeof adapter.exists === "function") {
          const exists = await adapter.exists(currentDir);
          if (!exists) continue;
        }
        const listed = await adapter.list(currentDir);
        for (const f of listed.files || []) {
          const norm = f.replace(/^\/+/, "");
          if (!this.isIgnored(norm)) {
            result.push(norm);
          }
        }
        for (const sub of listed.folders || []) {
          const normSub = sub.replace(/^\/+/, "");
          if (!this.isIgnored(normSub)) {
            queue.push(normSub);
          }
        }
      } catch {
        // Ignore unreadable config subdirectories
      }
    }

    return result;
  }

  private async localPathExists(relPath: string): Promise<boolean> {
    if (this.vault.getAbstractFileByPath(relPath)) {
      return true;
    }
    const adapter = (this.vault as unknown as {
      adapter?: { exists?: (p: string) => Promise<boolean> };
    }).adapter;
    if (adapter && typeof adapter.exists === "function") {
      return await adapter.exists(relPath);
    }
    return false;
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
        if (this.isConfigDirPath(trackedPath) && !this.shouldSyncObsidianConfig()) {
          continue;
        }
        const localExists = await this.localPathExists(trackedPath);
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
        if (this.isConfigDirPath(trackedFolder) && !this.shouldSyncObsidianConfig()) {
          continue;
        }
        const localFolder = await this.localPathExists(trackedFolder);
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

      // Also sync vault-specific settings inside .obsidian/ when enabled
      const configFiles = await this.listSyncableConfigFiles();
      for (const cfgPath of configFiles) {
        const pushStats = await this.pushPath(cfgPath);
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
    if (this.suppressDeleteForPaths.has(relPath) || this.isIgnored(relPath)) {
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
    if (this.isIgnored(newAbstract.path)) return;
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

        if (trackedFolderPath && !this.isIgnored(trackedFolderPath)) {
          const prefix = `${trackedFolderPath}/`;
          for (const trackedFile of Object.keys({ ...this.store.files })) {
            if (trackedFile.startsWith(prefix)) {
              await this.removePathSuppressed(trackedFile, false);
              delete this.store.files[trackedFile];
            }
          }
          for (const subFolder of Object.keys({ ...this.store.folders })) {
            if (subFolder === trackedFolderPath || subFolder.startsWith(prefix)) {
              delete this.store.folders[subFolder];
            }
          }
          await this.removePathSuppressed(trackedFolderPath, true);
          pulled++;
          continue;
        }

        const relPath =
          this.findTrackedFilePathByItemId(item.id) || this.resolveRelativePath(item);
        if (relPath && !this.isIgnored(relPath)) {
          await this.removePathSuppressed(relPath, false);
          delete this.store.files[relPath];
          pulled++;
        }
        continue;
      }

      const relPath = this.resolveRelativePath(item);
      if (!relPath || this.isIgnored(relPath)) {
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
        const existsLocally = await this.localPathExists(relPath);

        if (existsLocally) {
          const localBuf = await this.readBinaryByPath(relPath);
          const localHash = hashBinaryBuffer(localBuf);

          if (
            existingState?.binaryHash &&
            localHash !== existingState.binaryHash &&
            localHash !== remoteHash
          ) {
            // Concurrent binary edit: preserve local binary attachment as a conflict copy before writing remote
            const conflictPath = this.buildBinaryConflictPath(relPath);
            await this.writeBinaryByPathSuppressed(conflictPath, localBuf);
            conflicts++;
          }

          if (localHash !== remoteHash) {
            await this.writeBinaryByPathSuppressed(relPath, remoteBin.data);
          }
          this.store.files[relPath] = {
            itemId: item.id,
            eTag: remoteBin.eTag || item.eTag || "",
            baseText: "",
            isBinary: true,
            binaryHash: remoteHash,
            lastSyncedMtime: Date.now(),
            lastModifiedBy: remoteBin.lastModifiedBy,
          };
          pulled++;
        } else {
          await this.ensureParentFolders(relPath);
          await this.writeBinaryByPathSuppressed(relPath, remoteBin.data);
          this.store.files[relPath] = {
            itemId: item.id,
            eTag: remoteBin.eTag || item.eTag || "",
            baseText: "",
            isBinary: true,
            binaryHash: remoteHash,
            lastSyncedMtime: Date.now(),
            lastModifiedBy: remoteBin.lastModifiedBy,
          };
          pulled++;
        }
        continue;
      }

      // Text / Markdown / JSON / CSS file with 3-way diff3 merge
      const remote = await this.graph.downloadFileText(this.target.driveId, item.id);
      const existsLocally = await this.localPathExists(relPath);

      if (existsLocally) {
        const localText = await this.readTextByPath(relPath);

        if (!existingState || localText === existingState.baseText) {
          // Clean fast-forward update from colleague
          if (localText !== remote.text) {
            await this.writeTextByPathSuppressed(relPath, remote.text);
          }
          this.store.files[relPath] = {
            itemId: item.id,
            eTag: remote.eTag || item.eTag || "",
            baseText: remote.text,
            lastSyncedMtime: Date.now(),
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
          await this.writeTextByPathSuppressed(relPath, mergedText);

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
            lastSyncedMtime: Date.now(),
            lastModifiedBy: remote.lastModifiedBy,
          };
          merged++;
          if (hadConflict) conflicts++;
        }
      } else {
        // New file from SharePoint / OneDrive
        await this.ensureParentFolders(relPath);
        await this.writeTextByPathSuppressed(relPath, remote.text);
        this.store.files[relPath] = {
          itemId: item.id,
          eTag: remote.eTag || item.eTag || "",
          baseText: remote.text,
          lastSyncedMtime: Date.now(),
          lastModifiedBy: remote.lastModifiedBy,
        };
        pulled++;
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
    return this.pushPath(file.path);
  }

  /**
   * Push a file by relative path (supports both standard vault files and `.obsidian/` config files).
   */
  async pushPath(
    relPath: string
  ): Promise<{ pushed: number; merged: number; conflicts: number }> {
    if (this.suppressModifyForPaths.has(relPath) || this.isIgnored(relPath)) {
      return { pushed: 0, merged: 0, conflicts: 0 };
    }

    if (!isTextPath(relPath)) {
      return this.pushBinaryPath(relPath);
    }

    const localText = await this.readTextByPath(relPath);
    const state = this.store.files[relPath];

    if (state && !state.isBinary && localText === state.baseText) {
      return { pushed: 0, merged: 0, conflicts: 0 };
    }

    const res = await this.graph.uploadFileText(
      this.target,
      relPath,
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

      await this.writeTextByPathSuppressed(relPath, mergedText);
      const retry = await this.graph.uploadFileText(
        this.target,
        relPath,
        mergedText,
        remote.eTag
      );

      if (retry.item) {
        this.store.files[relPath] = {
          itemId: retry.item.id,
          eTag: retry.item.eTag ?? "",
          baseText: mergedText,
          lastSyncedMtime: Date.now(),
          lastModifiedBy: remote.lastModifiedBy,
        };
        await this.saveStore();
      }
      return { pushed: 1, merged: 1, conflicts: hadConflict ? 1 : 0 };
    }

    if (res.item) {
      this.store.files[relPath] = {
        itemId: res.item.id,
        eTag: res.item.eTag ?? "",
        baseText: localText,
        lastSyncedMtime: Date.now(),
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

  private async pushBinaryPath(
    relPath: string
  ): Promise<{ pushed: number; merged: number; conflicts: number }> {
    const localBuf = await this.readBinaryByPath(relPath);
    const localHash = hashBinaryBuffer(localBuf);
    const state = this.store.files[relPath];

    if (state?.isBinary && state.binaryHash === localHash) {
      return { pushed: 0, merged: 0, conflicts: 0 };
    }

    const res = await this.graph.uploadFileBinary(
      this.target,
      relPath,
      localBuf,
      state?.eTag
    );

    if (res.status === 412 && state?.itemId) {
      // Concurrent binary update: save local copy as conflict file and update main file to remote
      const remoteBin = await this.graph.downloadFileBinary(
        this.target.driveId,
        state.itemId
      );
      const conflictPath = this.buildBinaryConflictPath(relPath);
      await this.writeBinaryByPathSuppressed(conflictPath, localBuf);
      await this.writeBinaryByPathSuppressed(relPath, remoteBin.data);
      const remoteHash = hashBinaryBuffer(remoteBin.data);
      this.store.files[relPath] = {
        itemId: state.itemId,
        eTag: remoteBin.eTag,
        baseText: "",
        isBinary: true,
        binaryHash: remoteHash,
        lastSyncedMtime: Date.now(),
        lastModifiedBy: remoteBin.lastModifiedBy,
      };
      await this.pushBinaryPath(conflictPath);
      await this.saveStore();
      return { pushed: 1, merged: 0, conflicts: 1 };
    }

    if (res.item) {
      this.store.files[relPath] = {
        itemId: res.item.id,
        eTag: res.item.eTag ?? "",
        baseText: "",
        isBinary: true,
        binaryHash: localHash,
        lastSyncedMtime: Date.now(),
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

  private getAdapter(): {
    exists?: (p: string) => Promise<boolean>;
    read?: (p: string) => Promise<string>;
    readBinary?: (p: string) => Promise<ArrayBuffer>;
    write?: (p: string, d: string) => Promise<void>;
    writeBinary?: (p: string, d: ArrayBuffer) => Promise<void>;
    mkdir?: (p: string) => Promise<void>;
    remove?: (p: string) => Promise<void>;
    rmdir?: (p: string, recursive: boolean) => Promise<void>;
  } | undefined {
    return (this.vault as unknown as { adapter?: Record<string, unknown> }).adapter as
      | {
          exists?: (p: string) => Promise<boolean>;
          read?: (p: string) => Promise<string>;
          readBinary?: (p: string) => Promise<ArrayBuffer>;
          write?: (p: string, d: string) => Promise<void>;
          writeBinary?: (p: string, d: ArrayBuffer) => Promise<void>;
          mkdir?: (p: string) => Promise<void>;
          remove?: (p: string) => Promise<void>;
          rmdir?: (p: string, recursive: boolean) => Promise<void>;
        }
      | undefined;
  }

  private async readTextByPath(relPath: string): Promise<string> {
    const abstract = this.vault.getAbstractFileByPath(relPath);
    if (abstract instanceof TFile) {
      return await this.vault.read(abstract);
    }
    const adapter = this.getAdapter();
    if (adapter && typeof adapter.read === "function") {
      return await adapter.read(relPath);
    }
    return "";
  }

  private async readBinaryByPath(relPath: string): Promise<ArrayBuffer> {
    const abstract = this.vault.getAbstractFileByPath(relPath);
    if (abstract instanceof TFile) {
      const v = this.vault as unknown as {
        readBinary?: (f: TFile) => Promise<ArrayBuffer>;
      };
      if (typeof v.readBinary === "function") {
        return await v.readBinary(abstract);
      }
      const text = await this.vault.read(abstract);
      return new TextEncoder().encode(text).buffer;
    }
    const adapter = this.getAdapter();
    if (adapter && typeof adapter.readBinary === "function") {
      return await adapter.readBinary(relPath);
    }
    if (adapter && typeof adapter.read === "function") {
      const text = await adapter.read(relPath);
      return new TextEncoder().encode(text).buffer;
    }
    return new ArrayBuffer(0);
  }

  private async writeTextByPathSuppressed(
    relPath: string,
    content: string
  ): Promise<void> {
    this.suppressModifyForPaths.add(relPath);
    try {
      const abstract = this.vault.getAbstractFileByPath(relPath);
      if (abstract instanceof TFile) {
        await this.vault.modify(abstract, content);
        return;
      }
      const adapter = this.getAdapter();
      if (this.isConfigDirPath(relPath) && adapter && typeof adapter.write === "function") {
        await adapter.write(relPath, content);
        return;
      }
      await this.vault.create(relPath, content);
    } finally {
      this.suppressModifyForPaths.delete(relPath);
    }
  }

  private async writeBinaryByPathSuppressed(
    relPath: string,
    data: ArrayBuffer
  ): Promise<void> {
    this.suppressModifyForPaths.add(relPath);
    try {
      const abstract = this.vault.getAbstractFileByPath(relPath);
      const v = this.vault as unknown as {
        modifyBinary?: (f: TFile, d: ArrayBuffer) => Promise<void>;
        createBinary?: (p: string, d: ArrayBuffer) => Promise<TFile>;
      };
      if (abstract instanceof TFile) {
        if (typeof v.modifyBinary === "function") {
          await v.modifyBinary(abstract, data);
        } else {
          await this.vault.modify(abstract, new TextDecoder().decode(data));
        }
        return;
      }
      const adapter = this.getAdapter();
      if (
        this.isConfigDirPath(relPath) &&
        adapter &&
        typeof adapter.writeBinary === "function"
      ) {
        await adapter.writeBinary(relPath, data);
        return;
      }
      if (typeof v.createBinary === "function") {
        await v.createBinary(relPath, data);
      } else {
        await this.vault.create(relPath, new TextDecoder().decode(data));
      }
    } finally {
      this.suppressModifyForPaths.delete(relPath);
    }
  }

  private async removePathSuppressed(
    relPath: string,
    isFolder: boolean
  ): Promise<void> {
    this.suppressDeleteForPaths.add(relPath);
    try {
      const abstract = this.vault.getAbstractFileByPath(relPath);
      if (abstract) {
        await (
          this.vault as unknown as {
            delete: (f: unknown, force?: boolean) => Promise<void>;
          }
        ).delete(abstract, true);
        return;
      }
      const adapter = this.getAdapter();
      if (!adapter) return;
      if (typeof adapter.exists === "function" && !(await adapter.exists(relPath))) {
        return;
      }
      if (isFolder && typeof adapter.rmdir === "function") {
        await adapter.rmdir(relPath, true);
      } else if (typeof adapter.remove === "function") {
        await adapter.remove(relPath);
      }
    } finally {
      this.suppressDeleteForPaths.delete(relPath);
    }
  }

  private async ensureFolderExists(folderPath: string): Promise<void> {
    const parts = folderPath.split("/").filter(Boolean);
    let current = "";
    const adapter = this.getAdapter();
    for (const part of parts) {
      current = current ? `${current}/${part}` : part;
      if (this.isConfigDirPath(current) && adapter) {
        const exists =
          typeof adapter.exists === "function"
            ? await adapter.exists(current)
            : false;
        if (!exists && typeof adapter.mkdir === "function") {
          await adapter.mkdir(current);
        }
        continue;
      }
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
