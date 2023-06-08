import JSZip from 'jszip';
import { lookup } from 'mime-types';
import { basename } from 'path';
import { FileSystemUtilities } from '../utilities/file-system-utilities';
import logger from '../utilities/logger';

/**
 * Exposes read and write operations for the ooxml file.
 */
export class OOXMLPackageFileAccessor {
  private zip: JSZip | undefined;
  private get mimeType() {
    return lookup(basename(this.ooxmlPackagePath)) || undefined;
  }

  /**
   * Creates the ooxml package file accessor.
   *
   * @constructor
   * @param {string} ooxmlPackagePath The path to the ooxml package.
   */
  constructor(private ooxmlPackagePath: string) {}

  /**
   * Loads the ooxml package from the file system.
   */
  async load(): Promise<void> {
    logger.debug(`Loading ooxml package '${this.ooxmlPackagePath}'`);
    const data = FileSystemUtilities.readFile(this.ooxmlPackagePath);
    this.zip = new JSZip();
    await this.zip.loadAsync(data);
  }

  /**
   * Updates the ooxml package with the provided file.
   *
   * @param {string} filePath The path of the file inside the ooxml package to update.
   * @param {Uint8Array} data The new data for the file.
   * @returns {Promise<boolean>} True or false depending on whether the package updated successfully.
   */
  async updatePackage(filePath: string, data: Uint8Array): Promise<boolean> {
    logger.debug(`Updating '${filePath}' in OOXML package`);
    if (!this.zip) {
      logger.warn('Unable to update package since zip does not exist');
      return false;
    }

    const file = await this.zip.file(filePath, data).generateAsync({ type: 'uint8array', mimeType: this.mimeType, compression: 'DEFLATE' });
    return await FileSystemUtilities.writeFile(this.ooxmlPackagePath, file);
  }

  /**
   * Returns the contents of the ooxml package.
   *
   * @returns {PackageFile[]} The contents of the ooxml package.
   */
  getPackageContents(): Promise<PackageFile[]> {
    logger.debug('Getting OOXML package contents');
    if (!this.zip) {
      logger.warn('Unable to view package contents since zip does not exist');
      return Promise.resolve([]);
    }

    return Promise.all(
      Object.keys(this.zip.files)
        .sort()
        .map(async filePath => ({
          filePath: filePath,
          isDirectory: this.zip?.files[filePath].dir ?? false,
          data: (await this.zip?.file(filePath)?.async('uint8array')) ?? new Uint8Array(),
        })),
    );
  }
}

/**
 * Represents each file in the ooxml package.
 */
export interface PackageFile {
  isDirectory: boolean;
  filePath: string;
  data: Uint8Array;
}
