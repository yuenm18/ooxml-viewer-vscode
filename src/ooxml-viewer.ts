import { ExtensionContext, TreeView, Uri, window } from 'vscode';
import packageJson from '../package.json';
import { OOXMLPackage } from './ooxml-package';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';

const extensionName = packageJson.displayName;

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  treeView: TreeView<FileNode>;
  ooxmlPackages: OOXMLPackage[];

  /**
   * @description Constructs an instance of OOXMLViewer
   * @constructor OOXMLViewer
   * @param  {ExtensionContext} context The extension context
   * @returns {OOXMLViewer} instance
   */
  constructor(private context: ExtensionContext) {
    this.treeDataProvider = new OOXMLTreeDataProvider();
    this.treeView = window.createTreeView('ooxmlViewer', { treeDataProvider: this.treeDataProvider });
    this.treeView.title = extensionName;
    context.subscriptions.push(this.treeView);

    this.ooxmlPackages = [];
  }

  /**
   * @description Loads the selected OOXML file into the tree view
   * @method openOoxmlPackage
   * @async
   * @param {Uri} file The OOXML file
   * @returns {Promise<void>}
   */
  async openOoxmlPackage(file: Uri): Promise<void> {
    await this.removeOoxmlPackage(file);

    const ooxmlPackage = new OOXMLPackage(file, this.treeDataProvider, this.context);
    this.ooxmlPackages.push(ooxmlPackage);
  }

  /**
   * @description Resets the OOXML viewer
   * @method reset
   * @async
   * @returns {Promise<void>} Promise that returns void
   */
  async reset(): Promise<void> {
    await Promise.all(this.ooxmlPackages.map(ooxmlPackage => ooxmlPackage.reset()));
    this.ooxmlPackages = [];
  }

  /**
   * @description Removes an ooxml package if it exists
   * @method removeOoxmlPackage
   * @async
   * @private
   * @param {Uri} file The OOXML file
   * @returns {Promise<void>}
   */
  private async removeOoxmlPackage(file: Uri): Promise<void> {
    const existingPackageIndex = this.ooxmlPackages.findIndex(ooxmlPackage => ooxmlPackage.ooxmlFilePath === file.fsPath);
    if (existingPackageIndex !== -1) {
      const existingOoxmlPackage = this.ooxmlPackages.splice(existingPackageIndex, 1);
      await existingOoxmlPackage[0].reset();
    }
  }
}
