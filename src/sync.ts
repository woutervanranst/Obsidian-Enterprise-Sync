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

const FALLBACK_CONFIG_DIR = ["", "obsidian"].join(".");

export function isIgnoredVaultPath(
  relPath: string,
  syncObsidianConfig = false,
  configDir = FALLBACK_CONFIG_DIR
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

  const cleanConfigDir = (configDir || FALLBACK_CONFIG_DIR).replace(/^\/+|\/+$/g, "");
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
    return (this.vault.configDir || FALLBACK_CONFIG_DIR).replace(/^\/+|\/+$/g, "");
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
            const preserved = await this.deleteRemoteFileSafely(trackedPath, state);
            if (preserved) {
              summary.conflicts += 1;
              continue;
            }
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
          if (this.store.folders) {
            delete this.store.folders[trackedFolder];
          }
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
            const preserved = await this.deleteRemoteFileSafely(trackedPath, state);
            if (!preserved) {
              delete this.store.files[trackedPath];
              deletedCount++;
            }
          } else {
            delete this.store.files[trackedPath];
          }
        }
      }
      const anyRemainingChildren = Object.keys(this.store.files).some(
        (p) => p === relPath || p.startsWith(folderPrefix)
      );
      for (const [trackedFolder, id] of Object.entries({ ...this.store.folders })) {
        if (trackedFolder === relPath || trackedFolder.startsWith(folderPrefix)) {
          if (!anyRemainingChildren && id && trackedFolder !== relPath) {
            await this.graph.deleteItem(this.target.driveId, id);
          }
          delete this.store.folders[trackedFolder];
        }
      }
      if (!anyRemainingChildren) {
        if (folderItemId) {
          await this.graph.deleteItem(this.target.driveId, folderItemId);
        } else {
          await this.graph.deleteItemByPath(this.target, relPath);
        }
        deletedCount++;
      }
    } else {
      const state = this.store.files[relPath];
      if (state?.itemId) {
        const preserved = await this.deleteRemoteFileSafely(relPath, state);
        if (!preserved) {
          delete this.store.files[relPath];
          deletedCount++;
        }
      } else {
        await this.graph.deleteItemByPath(this.target, relPath);
      }
    }

    await this.saveStore();
    return deletedCount;
  }

  /**
   * Deletes a tracked remote file using If-Match optimistic concurrency.
   * If a collaborator modified the file remotely since our last sync (HTTP 412 with changed content),
   * restores the collaborator's newer version locally instead of destroying their work.
   * Returns true if a remote edit was detected and restored locally.
   */
  private async deleteRemoteFileSafely(
    relPath: string,
    state: SyncFileState
  ): Promise<boolean> {
    const delRes = await this.graph.deleteItemWithETag(
      this.target.driveId,
      state.itemId,
      state.eTag || undefined
    );
    if (!delRes.conflict) {
      return false;
    }

    // ETag changed on remote: verify whether content actually changed (vs metadata-only ETag bump)
    if (state.isBinary || !isTextPath(relPath)) {
      const remoteBin = await this.graph.downloadFileBinary(
        this.target.driveId,
        state.itemId
      );
      const remoteHash = hashBinaryBuffer(remoteBin.data);
      if (state.binaryHash && remoteHash === state.binaryHash) {
        await this.graph.deleteItem(this.target.driveId, state.itemId);
        return false;
      }
      this.logger.warn(
        `Delete/Modify conflict on '${relPath}': remote binary was updated by colleague — restoring remote version`
      );
      await this.ensureParentFolders(relPath);
      await this.writeBinaryByPathSuppressed(relPath, remoteBin.data);
      this.store.files[relPath] = {
        itemId: state.itemId,
        eTag: remoteBin.eTag,
        baseText: "",
        isBinary: true,
        binaryHash: remoteHash,
        lastSyncedMtime: Date.now(),
        lastModifiedBy: remoteBin.lastModifiedBy,
      };
      return true;
    }

    const remote = await this.graph.downloadFileText(
      this.target.driveId,
      state.itemId
    );
    if (remote.text === state.baseText) {
      await this.graph.deleteItem(this.target.driveId, state.itemId);
      return false;
    }

    this.logger.warn(
      `Delete/Modify conflict on '${relPath}': remote note was updated by colleague — restoring remote version`
    );
    await this.ensureParentFolders(relPath);
    await this.writeTextByPathSuppressed(relPath, remote.text);
    this.store.files[relPath] = {
      itemId: state.itemId,
      eTag: remote.eTag,
      baseText: remote.text,
      lastSyncedMtime: Date.now(),
      lastModifiedBy: remote.lastModifiedBy,
    };
    return true;
  }

  /**
   * Handle a local file or folder rename/move in Obsidian.
   */
  async renameLocalPath(
    oldPath: string,
    newAbstract: TFile | TFolder
  ): Promise<void> {
    const isFolder = newAbstract instanceof TFolder;
    if (!this.isIgnored(oldPath)) {
      await this.deleteLocalPath(oldPath, isFolder);
    }
    if (this.isIgnored(newAbstract.path)) return;

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

  private async hasUnsyncedLocalChanges(relPath: string): Promise<boolean> {
    const exists = await this.localPathExists(relPath);
    if (!exists) return false;
    const state = this.store.files[relPath];
    if (!state) return true;
    if (state.isBinary || !isTextPath(relPath)) {
      const localBuf = await this.readBinaryByPath(relPath);
      return hashBinaryBuffer(localBuf) !== state.binaryHash;
    }
    const localText = await this.readTextByPath(relPath);
    return localText !== state.baseText;
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

    // Real Microsoft Graph /delta tombstones omit `name` and `parentReference.path`,
    // providing only `id` and `parentReference.id`. Propagate known child paths up
    // through `parentReference.id` so nested deleted folders resolve regardless of order.
    let inferred = true;
    while (inferred) {
      inferred = false;
      for (const item of items) {
        if (
          !item.deleted ||
          !item.parentReference?.id ||
          item.parentReference.id === this.target.rootItemId
        ) {
          continue;
        }
        const knownPath =
          this.findTrackedFilePathByItemId(item.id) ||
          this.findTrackedFolderPathByItemId(item.id);
        if (!knownPath || !knownPath.includes("/")) continue;
        const parentFolderPath = knownPath.slice(0, knownPath.lastIndexOf("/"));
        if (parentFolderPath && !this.store.folders[parentFolderPath]) {
          this.store.folders[parentFolderPath] = item.parentReference.id;
          inferred = true;
        }
      }
    }

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
          let preservedAnyChild = false;
          for (const trackedFile of Object.keys({ ...this.store.files })) {
            if (trackedFile.startsWith(prefix)) {
              if (await this.hasUnsyncedLocalChanges(trackedFile)) {
                this.logger.warn(
                  `Modify/Delete conflict: folder '${trackedFolderPath}' was deleted remotely, preserving modified local file '${trackedFile}'`
                );
                delete this.store.files[trackedFile];
                conflicts++;
                preservedAnyChild = true;
              } else {
                await this.removePathSuppressed(trackedFile, false);
                delete this.store.files[trackedFile];
              }
            }
          }
          if (
            !preservedAnyChild &&
            this.getSyncableVaultFiles().some((f) => f.path.startsWith(prefix))
          ) {
            preservedAnyChild = true;
          }
          for (const subFolder of Object.keys({ ...this.store.folders })) {
            if (subFolder === trackedFolderPath || subFolder.startsWith(prefix)) {
              delete this.store.folders[subFolder];
            }
          }
          if (!preservedAnyChild) {
            await this.removePathSuppressed(trackedFolderPath, true);
            pulled++;
          }
          continue;
        }

        const relPath =
          this.findTrackedFilePathByItemId(item.id) || this.resolveRelativePath(item);
        if (relPath && !this.isIgnored(relPath)) {
          if (await this.hasUnsyncedLocalChanges(relPath)) {
            this.logger.warn(
              `Modify/Delete conflict: '${relPath}' was deleted remotely, preserving unsynced local edits`
            );
            delete this.store.files[relPath];
            conflicts++;
          } else {
            await this.removePathSuppressed(relPath, false);
            delete this.store.files[relPath];
            pulled++;
          }
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
          const localChangedSinceSync =
            !existingState ||
            !existingState.binaryHash ||
            localHash !== existingState.binaryHash;

          if (localChangedSinceSync && localHash !== remoteHash) {
            // Concurrent or Add/Add binary edit: preserve local binary attachment as a conflict copy before writing remote
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
          if (existingState) {
            this.logger.warn(
              `Delete/Modify conflict on '${relPath}': deleted locally offline while modified remotely — restoring remote binary`
            );
            conflicts++;
          }
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

        if (
          (existingState && localText === existingState.baseText) ||
          localText === remote.text
        ) {
          // Clean fast-forward update from colleague (or identical edits on both sides)
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
        } else {
          // Concurrent edit since baseText OR Add/Add conflict (!existingState -> empty base)
          const base = existingState ? existingState.baseText : "";
          let { text: mergedText, hadConflict } = this.threeWayMerge(
            base,
            localText,
            remote.text
          );
          await this.writeTextByPathSuppressed(relPath, mergedText);

          let uploadRes = await this.graph.uploadFileText(
            this.target,
            relPath,
            mergedText,
            remote.eTag
          );
          let finalAuthor = remote.lastModifiedBy;
          let fallbackETag = remote.eTag;

          if (uploadRes.status === 412) {
            const latestRemote = await this.graph.downloadFileText(
              this.target.driveId,
              item.id
            );
            const secondMerge = this.threeWayMerge(
              remote.text,
              mergedText,
              latestRemote.text
            );
            mergedText = secondMerge.text;
            hadConflict = hadConflict || secondMerge.hadConflict;
            finalAuthor = latestRemote.lastModifiedBy;
            fallbackETag = latestRemote.eTag;
            await this.writeTextByPathSuppressed(relPath, mergedText);
            uploadRes = await this.graph.uploadFileText(
              this.target,
              relPath,
              mergedText,
              latestRemote.eTag
            );
          }

          const finalItem = uploadRes.item ?? item;
          this.recordImmediateParentFolder(relPath, finalItem);
          this.store.files[relPath] = {
            itemId: finalItem.id,
            eTag: finalItem.eTag ?? fallbackETag,
            baseText: mergedText,
            lastSyncedMtime: Date.now(),
            lastModifiedBy: finalAuthor,
          };
          merged++;
          if (hadConflict) conflicts++;
        }
      } else {
        // New file from SharePoint / OneDrive (or deleted locally offline while modified remotely)
        if (existingState) {
          this.logger.warn(
            `Delete/Modify conflict on '${relPath}': deleted locally offline while modified remotely — restoring remote note`
          );
          conflicts++;
        }
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

    // Guard live untracked file creation against an unseen remote file at the same path (Add/Add conflict)
    if (!state && !this.isSyncing) {
      const existingRemote = await this.graph.getItemByPath(this.target, relPath);
      if (existingRemote && !existingRemote.folder) {
        const remote = await this.graph.downloadFileText(
          this.target.driveId,
          existingRemote.id
        );
        if (localText === remote.text) {
          this.recordImmediateParentFolder(relPath, existingRemote);
          this.store.files[relPath] = {
            itemId: existingRemote.id,
            eTag: remote.eTag || existingRemote.eTag || "",
            baseText: localText,
            lastSyncedMtime: Date.now(),
            lastModifiedBy: remote.lastModifiedBy,
          };
          await this.saveStore();
          return { pushed: 0, merged: 0, conflicts: 0 };
        }
        const { text: mergedText, hadConflict } = this.threeWayMerge(
          "",
          localText,
          remote.text
        );
        await this.writeTextByPathSuppressed(relPath, mergedText);
        const mergedUpload = await this.graph.uploadFileText(
          this.target,
          relPath,
          mergedText,
          remote.eTag || existingRemote.eTag
        );
        if (mergedUpload.item) {
          this.recordImmediateParentFolder(relPath, mergedUpload.item);
          this.store.files[relPath] = {
            itemId: mergedUpload.item.id,
            eTag: mergedUpload.item.eTag ?? "",
            baseText: mergedText,
            lastSyncedMtime: Date.now(),
            lastModifiedBy: remote.lastModifiedBy,
          };
          await this.saveStore();
        }
        return { pushed: 1, merged: 1, conflicts: hadConflict ? 1 : 0 };
      }
    }

    const res = await this.graph.uploadFileText(
      this.target,
      relPath,
      localText,
      state?.eTag
    );

    if (res.status === 412 && state?.itemId) {
      // Concurrent edit (or remote deletion) detected!
      let remote: { text: string; eTag: string; lastModifiedBy?: string } | null = null;
      try {
        remote = await this.graph.downloadFileText(this.target.driveId, state.itemId);
      } catch {
        remote = null;
      }

      if (!remote) {
        // Remote file was deleted while modified locally -> resurrect local version
        this.logger.warn(
          `Modify/Delete conflict on push for '${relPath}': remote item was deleted — re-creating with local edits`
        );
        const recreated = await this.graph.uploadFileText(
          this.target,
          relPath,
          localText
        );
        if (recreated.item) {
          this.recordImmediateParentFolder(relPath, recreated.item);
          this.store.files[relPath] = {
            itemId: recreated.item.id,
            eTag: recreated.item.eTag ?? "",
            baseText: localText,
            lastSyncedMtime: Date.now(),
          };
          await this.saveStore();
        }
        return { pushed: 1, merged: 0, conflicts: 1 };
      }

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
        this.recordImmediateParentFolder(relPath, retry.item);
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
      this.recordImmediateParentFolder(relPath, res.item);
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

    // Guard live untracked binary creation against an unseen remote binary at the same path (Add/Add conflict)
    if (!state && !this.isSyncing) {
      const existingRemote = await this.graph.getItemByPath(this.target, relPath);
      if (existingRemote && !existingRemote.folder) {
        const remoteBin = await this.graph.downloadFileBinary(
          this.target.driveId,
          existingRemote.id
        );
        const remoteHash = hashBinaryBuffer(remoteBin.data);
        if (localHash === remoteHash) {
          this.recordImmediateParentFolder(relPath, existingRemote);
          this.store.files[relPath] = {
            itemId: existingRemote.id,
            eTag: remoteBin.eTag || existingRemote.eTag || "",
            baseText: "",
            isBinary: true,
            binaryHash: remoteHash,
            lastSyncedMtime: Date.now(),
            lastModifiedBy: remoteBin.lastModifiedBy,
          };
          await this.saveStore();
          return { pushed: 0, merged: 0, conflicts: 0 };
        }
        const conflictPath = this.buildBinaryConflictPath(relPath);
        await this.writeBinaryByPathSuppressed(conflictPath, localBuf);
        await this.writeBinaryByPathSuppressed(relPath, remoteBin.data);
        this.recordImmediateParentFolder(relPath, existingRemote);
        this.store.files[relPath] = {
          itemId: existingRemote.id,
          eTag: remoteBin.eTag || existingRemote.eTag || "",
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
    }

    const res = await this.graph.uploadFileBinary(
      this.target,
      relPath,
      localBuf,
      state?.eTag
    );

    if (res.status === 412 && state?.itemId) {
      // Concurrent binary update (or remote deletion):
      let remoteBin: { data: ArrayBuffer; eTag: string; lastModifiedBy?: string } | null =
        null;
      try {
        remoteBin = await this.graph.downloadFileBinary(
          this.target.driveId,
          state.itemId
        );
      } catch {
        remoteBin = null;
      }

      if (!remoteBin) {
        const recreated = await this.graph.uploadFileBinary(
          this.target,
          relPath,
          localBuf
        );
        if (recreated.item) {
          this.recordImmediateParentFolder(relPath, recreated.item);
          this.store.files[relPath] = {
            itemId: recreated.item.id,
            eTag: recreated.item.eTag ?? "",
            baseText: "",
            isBinary: true,
            binaryHash: localHash,
            lastSyncedMtime: Date.now(),
          };
          await this.saveStore();
        }
        return { pushed: 1, merged: 0, conflicts: 1 };
      }

      // Save local copy as conflict file and update main file to remote
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
      this.recordImmediateParentFolder(relPath, res.item);
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
    return this.vault.adapter;
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
      const existing = this.vault.getAbstractFileByPath(current);
      if (existing instanceof TFolder) {
        continue;
      }
      if (existing) {
        throw new Error(`Cannot create folder '${current}': a local file already exists there.`);
      }
      if (
        adapter &&
        typeof adapter.exists === "function" &&
        await adapter.exists(current)
      ) {
        continue;
      }
      try {
        await this.vault.createFolder(current);
      } catch (err) {
        if (
          adapter &&
          typeof adapter.exists === "function" &&
          await adapter.exists(current)
        ) {
          continue;
        }
        throw err;
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

  private recordImmediateParentFolder(relPath: string, item?: DeltaDriveItem): void {
    if (!item?.parentReference?.id || item.parentReference.id === this.target.rootItemId) {
      return;
    }
    const slashIdx = relPath.lastIndexOf("/");
    if (slashIdx <= 0) return;
    const parentFolderPath = relPath.slice(0, slashIdx);
    if (!this.store.folders) {
      this.store.folders = {};
    }
    this.store.folders[parentFolderPath] = item.parentReference.id;
  }

  private resolveRelativePath(item: DeltaDriveItem): string {
    if (!item.name) {
      return "";
    }
    if (item.parentReference?.id && item.parentReference.id === this.target.rootItemId) {
      return item.name;
    }
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
    let afterColon = parentPath.slice(colonIdx + 1).replace(/^\/+/, "");
    try {
      afterColon = decodeURIComponent(afterColon);
    } catch {
      // Keep raw if not URI-encoded
    }
    if (!afterColon) {
      return item.name;
    }
    const segments = afterColon.split("/").filter(Boolean);
    if (this.target.rootItemId === "root") {
      return segments.length > 0 ? `${segments.join("/")}/${item.name}` : item.name;
    }
    // Strip the vault folder prefix if present
    const vaultFolderName = this.target.name.replace(/\s*\(.*?\)\s*$/, "").trim();
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
