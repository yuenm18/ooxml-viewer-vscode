import crypto from 'crypto';
import { dirname, join, sep } from 'path';
import { ExtensionUtilities } from '../utilities/extension-utilities';
import { FileSystemUtilities } from '../utilities/file-system-utilities';

const CACHE_FOLDER_NAME = 'cache';
const NORMAL_SUBFOLDER_NAME = 'normal';
const PREV_SUBFOLDER_NAME = 'prev';
const COMPARE_SUBFOLDER_NAME = 'compare';

/**
 * The file system cache for ooxml files.
 *
 * Normal cache files
 *  - used for displaying formatted files, editing ooxml parts, and as a compare point during a diff
 * Prev cache files
 *  - used as a compare point during saving to determine if the ooxml file should be resaved
 *  - these should always be updated with the normal cache file unless the ooxml parts are being updated
 * Compare cache files
 *  - used as a comparison point for diffs
 */
export class OOXMLPackageFileCache {
  private uniqueFileHash: string;
  /**
   * The base path where are files are stored in the cache.
   *
   * @returns {string} The base path of the cache.
   */
  private get cacheBasePath(): string {
    return join(this.storagePath, CACHE_FOLDER_NAME, this.uniqueFileHash);
  }

  get normalSubfolderPath(): string {
    return join(this.cacheBasePath, NORMAL_SUBFOLDER_NAME);
  }

  /**
   * Creates a new instance of an ooxml file cache.
   *
   * @constructor
   * @param {string} filePath The relative path to the ooxml file.
   * @param {string} storagePath The path to workspace storage directory.
   */
  constructor(filePath: string, private storagePath: string) {
    this.uniqueFileHash = crypto.createHash('sha256').update(filePath).digest('hex');
  }

  /**
   * Caches a file and its prev and compare parts.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file to cache.
   * @returns {Promise<void>}
   */
  async createCachedFiles(filePath: string, fileContents: Uint8Array): Promise<void> {
    await Promise.all([
      this.cacheNormalFile(filePath, fileContents),
      this.cachePrevFile(filePath, fileContents),
      this.cacheCompareFile(filePath, fileContents),
    ]);
  }

  /**
   * Caches a file and its prev and saves an empty compare file.
   * This is for when a new file is created so that the diff shows an empty
   * original file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file to cache.
   * @returns {Promise<void>}
   */
  async createCachedFilesWithEmptyCompare(filePath: string, fileContents: Uint8Array): Promise<void> {
    await Promise.all([
      this.cacheNormalFile(filePath, fileContents),
      this.cachePrevFile(filePath, fileContents),
      this.cacheCompareFile(filePath, new Uint8Array()),
    ]);
  }

  /**
   * Updates the cache for a file.
   * The compare file is updated to match the previous normal file
   * so that the diff works.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {Uint8Array} updatedFileContents The contents of the file to cache.
   * @returns {Promise<void>}
   */
  async updateCachedFiles(filePath: string, updatedFileContents: Uint8Array): Promise<void> {
    const cachedNormalFile = await this.getCachedNormalFile(filePath);
    await Promise.all([
      this.cacheNormalFile(filePath, updatedFileContents),
      this.cachePrevFile(filePath, updatedFileContents),
      this.cacheCompareFile(filePath, cachedNormalFile),
    ]);
  }

  /**
   * Updates the cache for a file and its prev but not the compare.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {Uint8Array} updatedFileContents The contents of the file to cache.
   * @returns {Promise<void>}
   */
  async updateCachedFilesNoCompare(filePath: string, updatedFileContents: Uint8Array): Promise<void> {
    await Promise.all([this.cacheNormalFile(filePath, updatedFileContents), this.cachePrevFile(filePath, updatedFileContents)]);
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
    await Promise.all([this.deleteNormalCachedFile(filePath), this.deleteCachedPrevFile(filePath), this.deleteCachedCompareFile(filePath)]);
  }

  /**
   * Gets normal file path of the cached file given the file path in an ooxml file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {string} The file path of the cached file.
   */
  getNormalFileCachePath(filePath: string): string {
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
    return !!cacheFilePath && cacheFilePath.startsWith(this.normalSubfolderPath);
  }

  /**
   * Gets a cached normal file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<Uint8Array>} Promise resolving to the contents of the file.
   */
  async getCachedNormalFile(filePath: string): Promise<Uint8Array> {
    const cachePath = this.getNormalFileCachePath(filePath);
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
   * Initializes the cache.
   *
   * @returns {Promise<void>}
   */
  async initialize(): Promise<void> {
    await FileSystemUtilities.createDirectory(this.cacheBasePath);
  }

  /**
   * Resets the cache.
   *
   * @returns {Promise<void>}
   */
  async reset(): Promise<void> {
    await this.deleteFile(this.cacheBasePath, true);
  }

  /**
   * Caches a normal file.
   *
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file.
   * @returns {Promise<void>}
   */
  private async cacheNormalFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const cachePath = this.getNormalFileCachePath(filePath);
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
  private async deleteNormalCachedFile(filePath: string): Promise<void> {
    const cachePath = this.getNormalFileCachePath(filePath);
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
   * Create a file from a part of the zip file.
   *
   * @param {string} cachedFilePath The path to the cached file.
   * @param {boolean} fileContents The file contents.
   * @returns {Promise<void>}
   */
  private async writeFile(cachedFilePath: string, fileContents: Uint8Array): Promise<void> {
    try {
      await FileSystemUtilities.createDirectory(dirname(this.cacheBasePath));
      await FileSystemUtilities.writeFile(cachedFilePath, fileContents);
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }
  }

  /**
   * Deletes a file in the cache.
   *
   * @param {string} cachedFilePath The path to the cached file.
   * @param {string} silentlyFail If true, swallow the exception and handle the error.
   * @returns {Promise<void>}
   */
  private async deleteFile(cachedFilePath: string, silentlyFail: boolean = false): Promise<void> {
    try {
      await FileSystemUtilities.deleteFile(cachedFilePath);
    } catch (err) {
      if (!silentlyFail) {
        await ExtensionUtilities.showError(err);
      }
    }
  }

  /**
   * Reads a file in the cache.
   *
   * @param {string} cachedFilePath The path to the cached file.
   * @returns {Promise<Uint8Array>} A promise resolving to the file contents.
   */
  private async readFile(cachedFilePath: string): Promise<Uint8Array> {
    try {
      return await FileSystemUtilities.readFile(cachedFilePath);
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }

    return new Uint8Array();
  }
}
