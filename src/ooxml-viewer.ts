import { basename } from 'path';
import { ExtensionContext } from 'vscode';
import { OOXMLExtensionSettings } from './ooxml-extension-settings';
import { OOXMLPackageFacade } from './ooxml-package/ooxml-package-facade';
import { OOXMLTreeDataProvider } from './tree-view/ooxml-tree-view-provider';
import { ExtensionUtilities } from './utilities/extension-utilities';
import { FileSystemUtilities } from './utilities/file-system-utilities';
import logger from './utilities/logger';

/**
 * The OOXML Viewer.
 */
export class OOXMLViewer {
  private ooxmlPackages: OOXMLPackageFacade[];

  private get contextStorageUri() {
    return (
      this.context.storageUri?.fsPath ??
      (() => {
        throw new Error('Storage URI does not exist');
      })()
    );
  }

  /**
   * Constructs an instance of OOXMLViewer.
   *
   * @constructor
   * @param  {OOXMLTreeDataProvider} treeDataProvider The tree data provider.
   * @param  {OOXMLExtensionSettings} settings The extension settings.
   * @param  {ExtensionContext} context The extension context.
   */
  constructor(
    private treeDataProvider: OOXMLTreeDataProvider,
    private settings: OOXMLExtensionSettings,
    private context: ExtensionContext,
  ) {
    this.ooxmlPackages = [];
  }

  /**
   * Loads the selected OOXML file into the tree view
   *
   * @param {string} filePath The OOXML file path.
   */
  async openOOXMLPackage(filePath: string): Promise<void> {
    logger.info(`Opening '${filePath}'`);
    await this.removeOOXMLPackage(filePath);

    const fileSize = await FileSystemUtilities.getFileSize(filePath);
    if (fileSize > this.settings.maximumOOXMLFileSizeBytes) {
      ExtensionUtilities.showWarning(
        `'${basename(filePath)}' size of '${fileSize}' exceeds the max file size of '${this.settings.maximumOOXMLFileSizeBytes}' bytes`,
      );
      return;
    }

    const ooxmlPackage = OOXMLPackageFacade.create(filePath, this.treeDataProvider, this.contextStorageUri);
    this.ooxmlPackages.push(ooxmlPackage);
    await ooxmlPackage.openOOXMLPackage();
  }

  /**
   * Removes the selected OOXML file into the tree view.
   *
   * @param {string} filePath The OOXML file path.
   */
  public async removeOOXMLPackage(filePath: string): Promise<void> {
    logger.debug(`Remove OOXML package called on '${filePath}'`);
    const existingPackageIndex = this.ooxmlPackages.findIndex(ooxmlPackage => ooxmlPackage.ooxmlFilePath === filePath);
    if (existingPackageIndex !== -1) {
      logger.info(`Removing '${filePath}'`);
      const existingOOXMLPackage = this.ooxmlPackages.splice(existingPackageIndex, 1);
      await existingOOXMLPackage[0].dispose();
    }
  }

  /**
   * Displays and formats the selected file.
   *
   * @param {string} ooxmlPackagePath The path to the ooxml file.
   * @param {string} filePath The selected file node's file path.
   */
  async viewFile(ooxmlPackagePath: string, filePath: string): Promise<void> {
    logger.info(`Viewing '${filePath}' in '${ooxmlPackagePath}'`);
    const ooxmlPackage = this.findOOXMLPackage(ooxmlPackagePath);
    await ooxmlPackage?.viewFile(filePath);
  }

  /**
   * Opens a window showing the difference between the primary xml part and the compare xml part.
   *
   * @param {string} ooxmlPackagePath The path to the ooxml file.
   * @param {string} filePath the path of file to be diffed.
   */
  async getDiff(ooxmlPackagePath: string, filePath: string): Promise<void> {
    logger.info(`Getting the diff of '${filePath}' in '${ooxmlPackagePath}'`);
    const ooxmlPackage = this.findOOXMLPackage(ooxmlPackagePath);
    await ooxmlPackage?.getDiff(filePath);
  }

  /**
   * Search OOXML parts for a string and display the results in the search window.
   *
   * @param {string} ooxmlPackagePath The path to the ooxml file.
   */
  async searchOOXMLParts(ooxmlPackagePath: string): Promise<void> {
    logger.info(`Searching '${ooxmlPackagePath}'`);
    const ooxmlPackage = this.findOOXMLPackage(ooxmlPackagePath);
    await ooxmlPackage?.searchOOXMLParts();
  }

  /**
   * Resets the OOXML viewer.
   */
  async reset(): Promise<void> {
    logger.info('Resetting the OOXML viewer');
    await Promise.all(this.ooxmlPackages.map(ooxmlPackage => ooxmlPackage.dispose()));
    this.ooxmlPackages = [];
    this.treeDataProvider.rootFileNode.children.length = 0;
    this.treeDataProvider.refresh();
    await this.tryClearCache();
  }

  private async tryClearCache(): Promise<void> {
    try {
      await FileSystemUtilities.deleteFile(this.contextStorageUri);
    } catch {
      logger.debug('Failed to clear the cache.');
    }
  }

  private findOOXMLPackage(filePath: string): OOXMLPackageFacade | undefined {
    return this.ooxmlPackages.find(ooxmlPackage => ooxmlPackage.ooxmlFilePath === filePath);
  }
}
