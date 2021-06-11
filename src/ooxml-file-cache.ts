import { existsSync } from 'fs';
import { dirname, join, sep } from 'path';
import { ExtensionContext, Uri, window, workspace } from 'vscode';

export const CACHE_FOLDER_NAME = 'cache';
export const NORMAL_SUBFOLDER_NAME = 'normal';
export const PREV_SUBFOLDER_NAME = 'prev';
export const COMPARE_SUBFOLDER_NAME = 'compare';

/**
 * The file system cache for ooxml files.
 *
 * Normal cache files
 *  - used for displaying formatted files, editting ooxml parts, and as a compare point during a diff
 * Prev cache files
 *  - used as a compare point during saving to determine if the ooxml file should be resaved
 *  - these should always be updated with the normal cache file unless the ooxml parts are being updated
 * Compare cache files
 *  - used as a comparison point for diffs
 */
export class OOXMLFileCache {
  /**
   * The base path where are files are stored in the cache.
   *
   * @returns {string} The base path of the cache.
   */
  get cacheBasePath(): string {
    return join(this.context.storageUri?.fsPath || '', CACHE_FOLDER_NAME);
  }

  get normalSubfolderPath(): string {
    return join(this.cacheBasePath, NORMAL_SUBFOLDER_NAME);
  }

  /**
   * Creates a new instance of an ooxml file cache.
   *
   * @constructor
   * @param {ExtensionContext} context The extension context.
   */
  constructor(private context: ExtensionContext) {
    this.initializeCache();
  }

  /**
   * Caches a file and its prev and compare parts.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file to cache.
   * @param {boolean} createEmptyCompareFile Whether or not to create an empty compare file.
   * @returns {Promise<void>}
   */
  async createCachedFile(filePath: string, fileContents: Uint8Array, createEmptyCompareFile: boolean): Promise<void> {
    await Promise.all([
      this.cacheFile(filePath, fileContents),
      this.cachePrevFile(filePath, fileContents),
      this.cacheCompareFile(filePath, createEmptyCompareFile ? new Uint8Array() : fileContents),
    ]);
  }

  /**
   * Updates the cache for a file and its prev and compare parts.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {Uint8Array} updatedFileContents The contents of the file to cache.
   * @param {boolean} updateCompareFile Whether or not the compare file should be updated with the previous version of the cache
   * @returns {Promise<void>}
   */
  async updateCachedFile(filePath: string, updatedFileContents: Uint8Array, updateCompareFile: boolean): Promise<void> {
    await Promise.all([
      this.cacheFile(filePath, updatedFileContents),
      this.cachePrevFile(filePath, updatedFileContents),
      updateCompareFile ? this.cacheCompareFile(filePath, await this.getCachedFile(filePath)) : Promise.resolve(),
    ]);
  }

  /**
   * Updates the compare file in the file cache.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file.
   * @returns {Promise<void>}
   */
  async updateCompareFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    return this.cacheCompareFile(filePath, fileContents);
  }

  /**
   * Deletes all cached parts of a file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  async deleteCachedFiles(filePath: string): Promise<void> {
    await Promise.all([this.deleteCachedFile(filePath), this.deleteCachedPrevFile(filePath), this.deleteCachedCompareFile(filePath)]);
  }

  /**
   * Gets file path of the cached file given the file path in an ooxml file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {string} The file path of the cached file.
   */
  getFileCachePath(filePath: string): string {
    return join(this.normalSubfolderPath, filePath);
  }

  /**
   * Gets file path of the cached compare file given the file path in an ooxml file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {string} The file path of the cached compare file.
   */
  getCompareFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, COMPARE_SUBFOLDER_NAME, filePath);
  }

  /**
   * Gets file path from the cache file path.
   *
   * @param {string} fileCachePath The file path in the cache.
   * @returns {string} The file path in the ooxml document file.
   */
  getFilePathFromCacheFilePath(cachePath: string): string {
    if (this.pathBelongsToCache(cachePath)) {
      // trim off base path plus the trailing separator
      const trimmedPath = cachePath.substring(this.cacheBasePath.length + sep.length);
      // trim off the subfolder and join path using unix file separators
      const normalizedPath = trimmedPath.split(sep).slice(1).join('/');
      return normalizedPath;
    }

    return cachePath;
  }

  /**
   * Determines whether or not the path belongs to the cache.
   *
   * @param {string} filePath The file path.
   * @returns {boolean} Whether or not the path is in the cache.
   */
  pathBelongsToCache(filePath: string): boolean {
    return !!filePath && filePath.startsWith(this.cacheBasePath);
  }

  /**
   * Determines whether or not the cache file path is for a normal cache file (not prev or compare).
   *
   * @param {string} cacheFilePath The file path.
   * @returns {boolean} Whether or not the path is in the cache.
   */
  cachePathIsNormal(cacheFilePath: string): boolean {
    return !!cacheFilePath && cacheFilePath.startsWith(join(this.cacheBasePath, NORMAL_SUBFOLDER_NAME));
  }

  /**
   * Gets a cached file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<Uint8Array>} Promise resolving to the contents of the file.
   */
  async getCachedFile(filePath: string): Promise<Uint8Array> {
    const cachePath = this.getFileCachePath(filePath);
    return this.readFile(cachePath);
  }

  /**
   * Gets a cached prev file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<Uint8Array>} Promise resolving to the contents of the file.
   */
  async getCachedPrevFile(filePath: string): Promise<Uint8Array> {
    const cachePrevPath = this.getPrevFileCachePath(filePath);
    return this.readFile(cachePrevPath);
  }

  /**
   * Gets a cached compare file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<Uint8Array>} Promise resolving to the contents of the file.
   */
  async getCachedCompareFile(filePath: string): Promise<Uint8Array> {
    const cacheComparePath = this.getCompareFileCachePath(filePath);
    return this.readFile(cacheComparePath);
  }

  /**
   * Caches a file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file.
   * @returns {Promise<void>}
   */
  private async cacheFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const cachePath = this.getFileCachePath(filePath);
    return this.writeFile(cachePath, fileContents);
  }

  /**
   * Caches a file prev.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file.
   * @returns {Promise<void>}
   */
  private async cachePrevFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const cachePrevPath = this.getPrevFileCachePath(filePath);
    return this.writeFile(cachePrevPath, fileContents);
  }

  /**
   * Caches a file compare.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file.
   * @returns {Promise<void>}
   */
  private async cacheCompareFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const cacheComparePath = this.getCompareFileCachePath(filePath);
    return this.writeFile(cacheComparePath, fileContents);
  }

  /**
   * Deletes a cached file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  private async deleteCachedFile(filePath: string): Promise<void> {
    const cachePath = this.getFileCachePath(filePath);
    return this.deleteFile(cachePath);
  }

  /**
   * Deletes a cached prev file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  private async deleteCachedPrevFile(filePath: string): Promise<void> {
    const cachePrevPath = this.getPrevFileCachePath(filePath);
    return this.deleteFile(cachePrevPath);
  }

  /**
   * Deletes a cached compare file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  private async deleteCachedCompareFile(filePath: string): Promise<void> {
    const cacheComparePath = this.getCompareFileCachePath(filePath);
    return this.deleteFile(cacheComparePath);
  }

  /**
   * Gets file path of the cached prev file given the file path in an ooxml file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {string} The file path of the cached prev file.
   */
  private getPrevFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, PREV_SUBFOLDER_NAME, filePath);
  }

  /**
   * Clears the cache.
   *
   * @returns {Promise<void>}
   */
  async clear(): Promise<void> {
    if (existsSync(this.cacheBasePath)) {
      await this.deleteFile(this.cacheBasePath);
      await this.initializeCache();
    }
  }

  /**
   * Create a file from a part of the zip file.
   *
   * @param {string} cachedFilePath The path to the cached file.
   * @param {boolean} fileContents The file contents.
   * @returns {Promise<void>}
   */
  async writeFile(cachedFilePath: string, fileContents: Uint8Array): Promise<void> {
    try {
      await workspace.fs.createDirectory(Uri.file(dirname(cachedFilePath)));
      await workspace.fs.writeFile(Uri.file(cachedFilePath), fileContents);
    } catch (err) {
      const msg = `Unable to create file '${cachedFilePath}\n ${err.message || err}`;
      console.error(msg);
      window.showErrorMessage(msg);
    }
  }

  /**
   * Deletes a file in the cache.
   *
   * @param {string} cachedFilePath The path to the cached file.
   * @returns {Promise<void>}
   */
  private async deleteFile(cachedFilePath: string): Promise<void> {
    try {
      await workspace.fs.delete(Uri.file(cachedFilePath), { recursive: true, useTrash: false });
    } catch (err) {
      const msg = `Unable to delete file '${cachedFilePath}\n ${err.message || err}`;
      console.error(msg);
      window.showErrorMessage(msg);
    }
  }

  /**
   * Reads a file in the cache.
   *
   * @param {string} cachedFilePath The path to the cached file.
   * @returns {Promise<Uint8Array>} A promise resolving to the file contents.
   */
  async readFile(cachedFilePath: string): Promise<Uint8Array> {
    try {
      return await workspace.fs.readFile(Uri.file(cachedFilePath));
    } catch (err) {
      const msg = `Unable to read file '${cachedFilePath}\n ${err.message || err}`;
      console.error(msg);
      window.showErrorMessage(msg);
    }

    return new Uint8Array();
  }

  /**
   * Initiailizes the cache.
   *
   * @returns {Promise<void>}
   */
  private async initializeCache(): Promise<void> {
    await workspace.fs.createDirectory(Uri.file(this.cacheBasePath));
  }
}
