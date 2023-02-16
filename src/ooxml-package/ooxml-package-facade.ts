import { getExtensionSettings } from '../ooxml-extension-settings';
import { OOXMLTreeDataProvider } from '../tree-view/ooxml-tree-view-provider';
import { OOXMLPackage } from './ooxml-package';
import { OOXMLPackageFileAccessor } from './ooxml-package-file-accessor';
import { OOXMLPackageFileCache } from './ooxml-package-file-cache';
import { OOXMLPackageFileWatcher } from './ooxml-package-file-watcher';
import { OOXMLPackageTreeView } from './ooxml-package-tree-view';

/**
 * Wraps all the functionality of the ooxml package.
 */
export class OOXMLPackageFacade {
  /**
   * Creates an instance of a OOXML package facade.
   *
   * @param filePath The path to the ooxml package.
   * @param treeDataProvider The tree data provider.
   * @param storagePath The path to the extension's storage path.
   * @returns {Promise<OOXMLPackageFacade>} The OOXML package facade.
   */
  static async create(filePath: string, treeDataProvider: OOXMLTreeDataProvider, storagePath: string): Promise<OOXMLPackageFacade> {
    const ooxmlFileCache = new OOXMLPackageFileCache(filePath, storagePath);
    const packageRootNode = new OOXMLPackageTreeView(treeDataProvider, filePath);
    const ooxmlFileAccessor = new OOXMLPackageFileAccessor(filePath);
    const ooxmlPackage = new OOXMLPackage(filePath, ooxmlFileAccessor, packageRootNode, ooxmlFileCache, getExtensionSettings());
    const fileWatchers = new OOXMLPackageFileWatcher(filePath, ooxmlPackage);

    await ooxmlPackage.openOOXMLPackage();

    return new OOXMLPackageFacade(filePath, ooxmlPackage, packageRootNode, fileWatchers, ooxmlFileCache);
  }

  private constructor(
    public ooxmlFilePath: string,
    private ooxmlPackage: OOXMLPackage,
    private packageRootNode: OOXMLPackageTreeView,
    private fileWatchers: OOXMLPackageFileWatcher,
    private fileCache: OOXMLPackageFileCache,
  ) {}

  /**
   * Displays and formats the selected file.
   *
   * @param {string} filePath The selected file node's file path.
   */
  async viewFile(filePath: string): Promise<void> {
    await this.ooxmlPackage.viewFile(filePath);
  }

  /**
   * Opens a window showing the difference between the primary xml part and the compare xml part.
   *
   * @param  {string} filePath The path of file to be diffed.
   */
  async getDiff(filePath: string): Promise<void> {
    await this.ooxmlPackage.getDiff(filePath);
  }

  /**
   * Search OOXML parts for a string and display the results in the search window.
   */
  async searchOOXMLParts(): Promise<void> {
    await this.ooxmlPackage.searchOOXMLParts();
  }

  /**
   * Disposes the ooxml package.
   */
  async dispose(): Promise<void> {
    this.fileWatchers.dispose();
    this.packageRootNode.reset();
    await this.fileCache.reset();
  }
}
