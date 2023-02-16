import { FileSystemError, Uri, workspace } from 'vscode';

/**
 * Wraps file system access.
 */
export class FileSystemUtilities {
  /**
   * Checks if a file exists.
   *
   * @param {string} filePath The file path to check.
   * @returns {Promise<boolean>} A promise that resolves to whether or not the file exists.
   */
  static async fileExists(filePath: string): Promise<boolean> {
    try {
      await workspace.fs.stat(Uri.file(filePath));
    } catch (err) {
      if ((err as FileSystemError)?.code?.toLowerCase() === 'filenotfound') {
        return false;
      }

      throw err;
    }

    return true;
  }

  /**
   * Gets the contents of the file.
   *
   * @param {string} filePath The path of the file to read.
   * @returns {Promise<Uint8Array>} A promise resolving to the contents of the file.
   */
  static async readFile(filePath: string): Promise<Uint8Array> {
    return await workspace.fs.readFile(Uri.file(filePath));
  }

  /**
   * Writes data to a file.
   *
   * @param filePath The path of the file to update.
   * @param data The data write to the file.
   * @returns A promise resolving to whether or not the file was written to successfully.
   */
  static async writeFile(filePath: string, data: Uint8Array): Promise<boolean> {
    try {
      await workspace.fs.writeFile(Uri.file(filePath), data);
    } catch (err) {
      if ((err as FileSystemError)?.code.toLowerCase() === 'unknown' && (err as FileSystemError)?.message.toLowerCase().includes('ebusy')) {
        return false;
      }

      throw err;
    }

    return true;
  }

  /**
   * Deletes a file.
   *
   * @param {string} filePath The path of the file to delete.
   */
  static async deleteFile(filePath: string): Promise<void> {
    await workspace.fs.delete(Uri.file(filePath), { recursive: true, useTrash: false });
  }

  /**
   * Creates a directory.
   *
   * @param {string} directoryPath The path to the directory to create.
   */
  static async createDirectory(directoryPath: string): Promise<void> {
    await workspace.fs.createDirectory(Uri.file(directoryPath));
  }
}
