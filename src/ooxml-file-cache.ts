import { existsSync } from 'fs';
import { basename, dirname, join } from 'path';
import { ExtensionContext, FileStat, Uri, workspace } from 'vscode';

export const CACHE_FOLDER_NAME = '.open-xml-viewer';

/**
 * The file system cache for ooxml files.
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
   * @param updatedFileContents The contents of the file to cache.
   * @returns {Promise<void>}
   */
  async updateCachedFile(filePath: string, updatedFileContents: Uint8Array): Promise<void> {
    const prevFileContents = await this.getCachedPrevFile(filePath);
    await Promise.all([
      this.cacheFile(filePath, updatedFileContents),
      this.cachePrevFile(filePath, updatedFileContents),
      this.cacheCompareFile(filePath, prevFileContents),
    ]);
  }

  /**
   * Deletes all cached parts of a file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  async deleteCachedFiles(filePath: string): Promise<void> {
    await Promise.all([
      this.deleteCachedFile(filePath),
      this.deleteCachedPrevFile(filePath),
      this.deleteCachedCompareFile(filePath),
    ]);
  }

  /**
   * Caches a file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file.
   * @returns {Promise<void>}
   */
  async cacheFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const cachePath = join(this.cacheBasePath, filePath);
    return this.writeFile(cachePath, fileContents);
  }

  /**
   * Caches a file prev.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @param {string} fileContents The contents of the file.
   * @returns {Promise<void>}
   */
  async cachePrevFile(filePath: string, fileContents: Uint8Array): Promise<void> {
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
  async cacheCompareFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const cacheComparePath = this.getCompareFileCachePath(filePath);
    return this.writeFile(cacheComparePath, fileContents);
  }
  
  /**
   * Deletes a cached file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  async deleteCachedFile(filePath: string): Promise<void> {
    const cachePath = join(this.cacheBasePath, filePath);
    return this.deleteFile(cachePath);
  }
  
  /**
   * Deletes a cached prev file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  async deleteCachedPrevFile(filePath: string): Promise<void> {
    const cachePrevPath = this.getPrevFileCachePath(filePath);
    return this.deleteFile(cachePrevPath);
  }
  
  /**
   * Deletes a cached compare file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<void>}
   */
  async deleteCachedCompareFile(filePath: string): Promise<void> {
    const cacheComparePath = this.getCompareFileCachePath(filePath);
    return this.deleteFile(cacheComparePath);
  }

  /**
   * Gets a cached file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<Uint8Array>} Promise resolving to the contents of the file.
   */
  async getCachedFile(filePath: string): Promise<Uint8Array> {
    const cachePath = join(this.cacheBasePath, filePath);
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
   * Gets file stats on a cached file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<FileStat>} Promise resolving to the stats of the file.
   */
  async getCachedFileStats(filePath: string): Promise<FileStat> {
    const cachePath = join(this.cacheBasePath, filePath);
    return this.getFileStats(cachePath);
  }

  /**
   * Gets file stats on the prev cached file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<FileStat>} Promise resolving to the stats of the file.
   */
  async getCachedPrevFileStats(filePath: string): Promise<FileStat> {
    const cachePrevPath = this.getPrevFileCachePath(filePath);
    return this.getFileStats(cachePrevPath);
  }

  /**
   * Gets file stats on the compare cached file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {Promise<FileStat>} Promise resolving to the stats of the file.
   */
  async getCachedCompareFileStats(filePath: string): Promise<FileStat> {
    const cacheComparePath = this.getCompareFileCachePath(filePath);
    return this.getFileStats(cacheComparePath);
  }

  /**
   * Gets file path of the cached file given the file path in an ooxml file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {string} The file path of the cached file.
   */
  getFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, filePath);
  }

  /**
   * Gets file path of the cached prev file given the file path in an ooxml file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {string} The file path of the cached prev file.
   */
  getPrevFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, this.getPrevFilePath(filePath));
  }

  /**
   * Gets file path of the cached compare file given the file path in an ooxml file.
   * 
   * @param {string} filePath The file path in the ooxml file.
   * @returns {string} The file path of the cached compare file.
   */
  getCompareFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, this.getCompareFilePath(filePath));
  }
  
  /**
   * Gets file path of the prev file given the path of the file in the cache.
   * 
   * @param {string} fileCachePath The file path in the cache.
   * @returns {string} The file path of the cached prev file.
   */
  getPrevFilePath(fileCachePath: string): string {
    return join(dirname(fileCachePath), `prev.${basename(fileCachePath)}`);
  }

  /**
   * Gets file path of the compare file given the path of the file in the cache.
   * 
   * @param {string} fileCachePath The file path in the cache.
   * @returns {string} The file path of the cached compare file.
   */
  getCompareFilePath(filePath: string): string {
    return join(dirname(filePath), `compare.${basename(filePath)}`);
  }

  /**
   * Clears the cache.
   * 
   * @returns {Promise<void>}
   */
  async clear(): Promise<void> {
    if (existsSync(this.cacheBasePath)) {
      await workspace.fs.delete(Uri.file(this.cacheBasePath), { recursive: true, useTrash: false });
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
      console.error(`Unable to create file '${cachedFilePath}'`, err);
    }
  }

  /**
   * Deletes a file in the cache.
   * 
   * @param {string} cachedFilePath The path to the cached file.
   * @returns {Promise<void>}
   */
  async deleteFile(cachedFilePath: string): Promise<void> {
    try {
      await workspace.fs.delete(Uri.file(cachedFilePath), { recursive: true, useTrash: false });
    } catch (err) {
      console.error(`Unable to delete file '${cachedFilePath}'`, err);
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
      console.error(`Unable to read file '${cachedFilePath}'`, err);
    }

    return new Uint8Array();
  }

  /**
   * Reads a file's stats in the cache.
   * 
   * @throws Throws if the file does not exist.
   * @param {string} cachedFilePath The path to the cached file.
   * @returns {Promise<FileStat>} A promise resolving to the file stats.
   */
  async getFileStats(cachedFilePath: string): Promise<FileStat> {
    try {
      return await workspace.fs.stat(Uri.file(cachedFilePath));
    } catch (err) {
      console.error(`Unable to read file '${cachedFilePath}' stats`, err);
      throw err;
    }
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