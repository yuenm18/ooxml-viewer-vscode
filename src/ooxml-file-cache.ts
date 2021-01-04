import { existsSync } from 'fs';
import { basename, dirname, join } from 'path';
import { ExtensionContext, FileStat, Uri, workspace } from 'vscode';

export const CACHE_FOLDER_NAME = '.open-xml-viewer';

export class OOXMLFileCache {
  get cacheBasePath(): string {
    return join(this._context.storageUri?.fsPath || '', CACHE_FOLDER_NAME);
  }

  constructor(private _context: ExtensionContext) {
    this.initializeCache();
  }

  async initializeCache(): Promise<void> {
    await workspace.fs.createDirectory(Uri.file(this.cacheBasePath));
  }

  // create a copy of the file and a prev copy with the same data i.e. no changes have been made
  async cacheCachedFile(filePath: string, fileContents: Uint8Array, createEmptyCompareFile: boolean): Promise<void> {
    await Promise.all([
      this.cacheFile(filePath, fileContents),
      this.cachePrevFile(filePath, fileContents),
      this.cacheCompareFile(filePath, createEmptyCompareFile ? new Uint8Array() : fileContents),
    ]);
  }

  async updateCachedFile(filePath: string, updatedFileContents: Uint8Array): Promise<void> {
    const prevFileContents = await this.getPrevFile(filePath);
    await Promise.all([
      this.cacheFile(filePath, updatedFileContents),
      this.cachePrevFile(filePath, updatedFileContents),
      this.cacheCompareFile(filePath, prevFileContents),
    ]);
  }

  async deleteCachedFiles(filePath: string): Promise<void> {
    await Promise.all([
      this.deleteCachedFile(filePath),
      this.deleteCachedPrevFile(filePath),
      this.deleteCachedCompareFile(filePath),
    ]);
  }

  async cacheFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    return this.createFile(filePath, fileContents);
  }

  async cachePrevFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const prevPath = join(dirname(filePath), `prev.${basename(filePath)}`);
    return this.createFile(prevPath, fileContents);
  }

  async cacheCompareFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    const comparePath = join(dirname(filePath), `compare.${basename(filePath)}`);
    return this.createFile(comparePath, fileContents);
  }
  
  async deleteCachedFile(filePath: string): Promise<void> {
    return this.deleteFile(filePath);
  }

  async deleteCachedPrevFile(filePath: string): Promise<void> {
    const prevPath = join(dirname(filePath), `prev.${basename(filePath)}`);
    return this.deleteFile(prevPath);
  }

  async deleteCachedCompareFile(filePath: string): Promise<void> {
    const comparePath = join(dirname(filePath), `compare.${basename(filePath)}`);
    return this.deleteFile(comparePath);
  }

  async getFile(filePath: string): Promise<Uint8Array> {
    return this.readFile(filePath);
  }

  async getPrevFile(filePath: string): Promise<Uint8Array> {
    const prevPath = join(dirname(filePath), `prev.${basename(filePath)}`);
    return this.readFile(prevPath);
  }

  async getCompareFile(filePath: string): Promise<Uint8Array> {
    const comparePath = join(dirname(filePath), `compare.${basename(filePath)}`);
    return this.readFile(comparePath);
  }

  async getFileStats(filePath: string): Promise<FileStat> {
    return this.readFileStats(filePath);
  }

  async getPrevFileStats(filePath: string): Promise<FileStat> {
    const prevPath = join(dirname(filePath), `prev.${basename(filePath)}`);
    return this.readFileStats(prevPath);
  }

  async getCompareFileStats(filePath: string): Promise<FileStat> {
    const comparePath = join(dirname(filePath), `compare.${basename(filePath)}`);
    return this.readFileStats(comparePath);
  }

  getFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, filePath);
  }

  getPrevFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, dirname(filePath), `prev.${basename(filePath)}`);
  }

  getCompareFileCachePath(filePath: string): string {
    return join(this.cacheBasePath, dirname(filePath), `compare.${basename(filePath)}`);
  }
  
  getPrevFilePath(filePath: string): string {
    return join(dirname(filePath), `prev.${basename(filePath)}`);
  }
  
  getCompareFilePath(filePath: string): string {
    return join(dirname(filePath), `compare.${basename(filePath)}`);
  }

  async clear(): Promise<void> {
    if (existsSync(this.cacheBasePath)) {
      return workspace.fs.delete(Uri.file(this.cacheBasePath), { recursive: true, useTrash: false });
    }
  }

  /**
   * @description Create a file from a part of the zip file
   * @method createFile
   * @async
   * @param  {string} relativePath The path to the folder in the zip file
   * @param  {string} fileName The name of the file
   * @param  {boolean} formatIt Whether or not the file should be formatted
   * @returns {Promise<void>}
   */
  async createFile(filePath: string, fileContents: Uint8Array): Promise<void> {
    try {
      const folderPath = join(this.cacheBasePath, dirname(filePath));
      const fileCachePath: string = join(folderPath, basename(filePath));
      await workspace.fs.createDirectory(Uri.file(folderPath));
      await workspace.fs.writeFile(Uri.file(fileCachePath), fileContents);
    } catch (err) {
      console.error(err);
    }
  }

  async deleteFile(filePath: string): Promise<void> {
    try {
      const fileCachePath = join(this.cacheBasePath, filePath);
      await workspace.fs.delete(Uri.file(fileCachePath), { recursive: true, useTrash: false });
    } catch (err) {
      console.error(err);
    }
  }

  async readFile(filePath: string): Promise<Uint8Array> {
    try {
      const fileCachePath = join(this.cacheBasePath, filePath);
      return await workspace.fs.readFile(Uri.file(fileCachePath));
    } catch (err) {
      console.error(err);
    }

    return new Uint8Array();
  }

  async readFileStats(filePath: string): Promise<FileStat> {
    try {
      const fileCachePath = join(this.cacheBasePath, filePath);
      return await workspace.fs.stat(Uri.file(fileCachePath));
    } catch (err) {
      console.error(err);
      throw err;
    }
  }
}