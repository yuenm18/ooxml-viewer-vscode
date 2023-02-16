import { ExtensionContext } from 'vscode';
import { OOXMLPackageFacade } from './ooxml-package/ooxml-package-facade';
import { OOXMLTreeDataProvider } from './tree-view/ooxml-tree-view-provider';
import { FileSystemUtilities } from './utilities/file-system-utilities';

/**
 * The OOXML Viewer.
 */
export class OOXMLViewer {
  private ooxmlPackages: OOXMLPackageFacade[];

  private get contextStorageUri() {
    return this.context.storageUri?.fsPath || '';
  }

  /**
   * Constructs an instance of OOXMLViewer.
   *
   * @constructor
   * @param  {OOXMLTreeDataProvider} treeDataProvider The tree data provider.
   * @param  {ExtensionContext} context The extension context.
   */
  constructor(private treeDataProvider: OOXMLTreeDataProvider, private context: ExtensionContext) {
    this.ooxmlPackages = [];
  }

  /**
   * Loads the selected OOXML file into the tree view
   *
   * @param {string} filePath The OOXML file path.
   */
  async openOOXMLPackage(filePath: string): Promise<void> {
    await this.removeOOXMLPackage(filePath);

    const ooxmlPackage = await OOXMLPackageFacade.create(filePath, this.treeDataProvider, this.contextStorageUri);
    this.ooxmlPackages.push(ooxmlPackage);
  }

  /**
   * Removes the selected OOXML file into the tree view.
   *
   * @param {string} filePath The OOXML file path.
   */
  public async removeOOXMLPackage(filePath: string): Promise<void> {
    const existingPackageIndex = this.ooxmlPackages.findIndex(ooxmlPackage => ooxmlPackage.ooxmlFilePath === filePath);
    if (existingPackageIndex !== -1) {
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
    const ooxmlPackage = this.findOOXMLPackage(ooxmlPackagePath);
    await ooxmlPackage?.getDiff(filePath);
  }

  /**
   * Search OOXML parts for a string and display the results in the search window.
   *
   * @param {string} ooxmlPackagePath The path to the ooxml file.
   */
  async searchOOXMLParts(ooxmlPackagePath: string): Promise<void> {
    const ooxmlPackage = this.findOOXMLPackage(ooxmlPackagePath);
    await ooxmlPackage?.searchOOXMLParts();
  }

  /**
   * Resets the OOXML viewer.
   */
  async reset(): Promise<void> {
    await Promise.all(this.ooxmlPackages.map(ooxmlPackage => ooxmlPackage.dispose()));
    this.ooxmlPackages = [];
    await this.tryClearCache();
  }

  private async tryClearCache(): Promise<void> {
    try {
      await FileSystemUtilities.deleteFile(this.contextStorageUri);
    } catch {}
  }

  private findOOXMLPackage(filePath: string): OOXMLPackageFacade | undefined {
    return this.ooxmlPackages.find(ooxmlPackage => ooxmlPackage.ooxmlFilePath === filePath);
  }
}
