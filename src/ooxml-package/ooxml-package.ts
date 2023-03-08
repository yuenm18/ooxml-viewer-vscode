import { basename } from 'path';
import { OOXMLExtensionSettings } from '../ooxml-extension-settings';
import { FileNode, FileNodeType } from '../tree-view/ooxml-tree-view-provider';
import { ExtensionUtilities } from '../utilities/extension-utilities';
import { RemoveOOXMLCommand } from '../utilities/ooxml-commands';
import { XmlFormatter } from '../utilities/xml-formatter';
import { OOXMLPackageFileAccessor } from './ooxml-package-file-accessor';
import { OOXMLPackageFileCache } from './ooxml-package-file-cache';
import { OOXMLPackageTreeView } from './ooxml-package-tree-view';

/**
 * The OOXML Package
 */
export class OOXMLPackage {
  // Whether or not this is the first time the file has been populated
  //  (Should the new file label (asterisk) be shown when creating a new file node).
  private isFirstOpen: boolean;
  private packageName: string;

  /**
   * Constructs an instance of OOXMLPackage.
   *
   * @constructor
   * @param {string} ooxmlFilePath The path to the ooxml file.
   * @param {OOXMLPackageFileAccessor} ooxmlFileAccessor The ooxml package file accessor.
   * @param {OOXMLPackageTreeView} treeView The package tree view.
   * @param {OOXMLPackageFileCache} cache The file cache for the ooxml package.
   * @param {OOXMLExtensionSettings} extensionSettings The extension settings.
   */
  constructor(
    private ooxmlFilePath: string,
    private ooxmlFileAccessor: OOXMLPackageFileAccessor,
    private treeView: OOXMLPackageTreeView,
    private cache: OOXMLPackageFileCache,
    private extensionSettings: OOXMLExtensionSettings,
  ) {
    this.isFirstOpen = true;
    this.packageName = basename(ooxmlFilePath);
  }

  /**
   * Displays and formats the selected file.
   *
   * @param {string} filePath The selected file node's file path
   */
  async viewFile(filePath: string): Promise<void> {
    try {
      await ExtensionUtilities.withProgress(async () => {
        await this.formatXml(filePath);

        const fileCachePath = this.cache.getNormalFileCachePath(filePath);
        await ExtensionUtilities.openFile(fileCachePath);
      }, `Opening ${filePath}`);
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }
  }

  /**
   * Opens a window showing the difference between the primary xml part and the compare xml part.
   *
   * @param {string} filePath The path of the file to be diffed.
   */
  async getDiff(filePath: string): Promise<void> {
    try {
      // format the file
      await this.formatXml(filePath);

      // diff the primary and compare files
      const fileCachePath = this.cache.getNormalFileCachePath(filePath);
      const fileCompareCachePath = this.cache.getCompareFileCachePath(filePath);
      const title = `${basename(fileCachePath)} ↔ compare.${basename(fileCompareCachePath)}`;

      await ExtensionUtilities.openDiff(fileCompareCachePath, fileCachePath, title);
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }
  }

  /**
   * Format the document if it is a cached normal file.
   *
   * @param {string} filePath The file path to format.
   */
  async tryFormatDocument(filePath: string): Promise<void> {
    try {
      if (this.cache.cachePathIsNormal(filePath)) {
        await ExtensionUtilities.withProgress(async () => {
          await this.formatXml(this.cache.getFilePathFromCacheFilePath(filePath));
        }, `Formatting '${basename(filePath)}'`);
      }
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }
  }

  /**
   * Search OOXML parts for a string and display the results in the search.
   */
  async searchOOXMLParts(): Promise<void> {
    try {
      const searchTerm = await ExtensionUtilities.showInput(`Search '${this.packageName}' OOXML Parts`, 'Enter a search term.');
      if (!searchTerm) {
        return;
      }

      await ExtensionUtilities.findInFiles(searchTerm, this.cache.normalSubfolderPath);
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }
  }

  /**
   * Loads the selected OOXML file into the tree view.
   */
  async openOOXMLPackage(): Promise<void> {
    try {
      await ExtensionUtilities.withProgress(async () => {
        // load ooxml file and populate the viewer
        await this.cache.initialize();
        await this.ooxmlFileAccessor.load();
        await this.populateOOXMLViewer();
      }, `Unpacking '${this.packageName}'`);
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }
  }

  /**
   * Writes changes to OOXML file being inspected when one of its parts is saved.
   * Note that this will trigger `reloadOOXMLFile` to fire if changes are written.
   *
   * @param {string} cacheFilePath The path of the xml part that was updated.
   */
  async updateOOXMLFile(cacheFilePath: string): Promise<void> {
    try {
      if (!this.cache.cachePathIsNormal(cacheFilePath)) {
        return;
      }

      const filePath = this.cache.getFilePathFromCacheFilePath(cacheFilePath);

      const fileContents = await this.cache.getCachedNormalFile(filePath);
      const prevFileContents = await this.cache.getCachedPrevFile(filePath);
      if (XmlFormatter.areEqual(fileContents, prevFileContents)) {
        return;
      }

      const fileMinXml = XmlFormatter.minify(fileContents, this.extensionSettings.preserveComments);

      const success = await this.ooxmlFileAccessor.updatePackage(filePath, fileMinXml);
      if (!success) {
        await ExtensionUtilities.showWarning(
          `File not saved.\n'${this.packageName}' is open in another program.\nClose that program before making any changes.`,
          true,
        );

        await ExtensionUtilities.makeActiveTextEditorDirty();
        return;
      }

      await this.cache.createCachedFiles(filePath, fileContents);

      this.treeView.refresh();
    } catch (err) {
      await ExtensionUtilities.showError(err);
    }
  }

  /**
   * Creates or updates tree view file nodes and creates cache files for comparison.
   */
  private async populateOOXMLViewer(): Promise<void> {
    const fileContents = await this.ooxmlFileAccessor.getPackageContents();

    const ooxmlContentsLength = fileContents.length;
    if (ooxmlContentsLength > this.extensionSettings.maximumNumberOfOOXMLParts) {
      ExtensionUtilities.showWarning(
        `'${this.packageName}' number of parts of '${ooxmlContentsLength}' exceeds the maximum of '${this.extensionSettings.maximumNumberOfOOXMLParts}'`,
      );
      await ExtensionUtilities.dispatch(new RemoveOOXMLCommand(this.treeView.getRootFileNode()));
      return;
    }

    for (const file of fileContents) {
      // ignore folder files
      if (file.isDirectory) {
        continue;
      }

      // Build nodes for each file in the package

      let fileNodeAlreadyExists = true;
      let currentFileNode = this.treeView.getRootFileNode();
      const names: string[] = file.filePath.split('/');
      for (let i = 0; i < names.length; i++) {
        const fileOrFolderPath = names.slice(0, i + 1).join('/');
        const existingFileNode = currentFileNode.children.find(c => c.nodePath === fileOrFolderPath);
        if (existingFileNode) {
          currentFileNode = existingFileNode;
        } else {
          fileNodeAlreadyExists = false;
          // create a new FileNode with the currentFileNode as parent and add it to the currentFileNode children
          const newFileNode = FileNode.create(fileOrFolderPath, currentFileNode, this.ooxmlFilePath);
          currentFileNode = newFileNode;
        }
      }

      // cache or update the cache of the node and mark the status of the node

      // If the file node already exists and it isn't already marked as deleted,
      // the next state of the node can either
      // - "modified" if the file has been changed from the outside (handled in this if block)
      // - "unchanged" if the file has not been changed from the outside (handled in this if block)
      // - "deleted" if the file isn't included in the new ooxml package (handled in handleDeletedParts())
      // If the file node does not already exist, the only possible next state is "created" (handled in else block)
      // If the file node exists and is marked as deleted already, the next possible states are
      // - no node if the file isn't included in the new ooxml package (handled in handleDeletedParts())
      // - "created" if the file is recreated (handled in the else block)

      if (fileNodeAlreadyExists && !currentFileNode.isDeleted()) {
        const filesAreDifferent = await this.hasFileBeenChangedFromOutside(currentFileNode.nodePath, file.data);
        await this.cache.updateCachedFiles(currentFileNode.nodePath, file.data);

        if (filesAreDifferent) {
          currentFileNode.setModified();
        } else {
          currentFileNode.setUnchanged();
        }
      } else {
        if (!this.isFirstOpen) {
          await this.cache.createCachedFilesWithEmptyCompare(currentFileNode.nodePath, file.data);
          currentFileNode.setCreated();
        } else {
          await this.cache.createCachedFiles(currentFileNode.nodePath, file.data);
        }
      }
    }

    // need to handle deleted parts separately since the zip
    // doesn't contain them anymore
    await this.handleDeletedParts(fileContents.map(file => file.filePath));
    await this.reformatOpenTabs(fileContents.map(file => file.filePath));

    // tell vscode the tree has changed
    this.treeView.refresh();

    this.isFirstOpen = false;
  }

  /**
   * Traverse tree and delete cached parts that don't exist anymore.
   * If the file node is marked as deleted, delete the cached part, if it isn't, then
   * mark the node for deletion on next update.
   *
   * @param {string[]} filePaths The file paths in the ooxml file.
   */
  private async handleDeletedParts(filePaths: string[]): Promise<void> {
    const filesInOOXMLFile = new Set(filePaths);
    const fileNodeQueue = [this.treeView.getRootFileNode()];

    let fileNode;
    while ((fileNode = fileNodeQueue.pop())) {
      if (fileNode.contextValue === FileNodeType.File && !filesInOOXMLFile.has(fileNode.nodePath)) {
        if (!fileNode.isDeleted()) {
          fileNode.setDeleted();
          await this.cache.updateCachedFiles(fileNode.nodePath, new Uint8Array());
        } else {
          // remove files marked as deleted from tree view and cache after the ooxml file
          // the second time the ooxml file is saved
          await this.cache.deleteCachedFiles(fileNode.nodePath);
          fileNode.parent?.children.splice(fileNode.parent.children.indexOf(fileNode), 1);
        }
      }

      fileNodeQueue.push(...fileNode.children);
    }

    this.treeView.refresh();
  }

  /**
   * Reformats the open tabs after their contents have been updated.
   * (Formats the xml of all tabs and closes the ones that were deleted)
   *
   * @param {string[]} filePaths The file paths in the ooxml file.
   */
  private async reformatOpenTabs(filePaths: string[]): Promise<void> {
    const filePathsInOOXMLPackage = new Set(filePaths);
    const openTextDocumentsInCache = ExtensionUtilities.getOpenTextDocumentFilePaths()
      .filter(fileName => this.cache.pathBelongsToCache(fileName))
      .map(fileName => this.cache.getFilePathFromCacheFilePath(fileName));

    const formatCacheDocuments = openTextDocumentsInCache
      .filter(p => filePathsInOOXMLPackage.has(p))
      .map(filePath => this.formatXml(filePath));

    const closeDocumentsInCacheButNotPackage = openTextDocumentsInCache
      .filter(fileName => !filePathsInOOXMLPackage.has(fileName))
      .map(fileName => ExtensionUtilities.closeTextDocument(fileName));

    await Promise.all([...formatCacheDocuments, ...closeDocumentsInCacheButNotPackage]);
  }

  /**
   * Tries to format the file as xml and if successful, updates the cached files.
   *
   * @param {string} filePath The path of the file in the ooxml package.
   */
  private async formatXml(filePath: string): Promise<void> {
    const data = await this.cache.getCachedNormalFile(filePath);
    const fileSize = XmlFormatter.minify(data, true).byteLength;
    if (fileSize > this.extensionSettings.maximumXmlPartsFileSizeBytes) {
      if (XmlFormatter.isXml(data)) {
        ExtensionUtilities.showWarning(
          `'${basename(filePath)}' size of '${fileSize}' exceeds maximum of '${this.extensionSettings.maximumXmlPartsFileSizeBytes}' bytes`,
        );
      }

      return;
    }

    // Need to format the normal/prev and the compare separately for the diff to work
    const formatNormalXml = async () => {
      const fileContent = await this.cache.getCachedNormalFile(filePath);
      const formattedXml = XmlFormatter.format(fileContent);

      if (fileContent != formattedXml) {
        await this.cache.updateCachedFilesNoCompare(filePath, formattedXml);
      }
    };

    const formatCompareXml = async () => {
      const compareFileContent = await this.cache.getCachedCompareFile(filePath);
      const formattedXml = XmlFormatter.format(compareFileContent);

      if (compareFileContent != formattedXml) {
        await this.cache.updateCompareFile(filePath, formattedXml);
      }
    };

    await Promise.all([formatNormalXml(), formatCompareXml()]);
  }

  /**
   * Check if an OOXML part is different from its cached version.
   *
   * @param {string} filePath The path of the file in the ooxml package.
   * @param {string} newContent The updated contents of the file.
   * @returns {Promise<boolean>} A Promise resolving to whether or not the file has been changed from the outside.
   */
  private async hasFileBeenChangedFromOutside(filePath: string, newContent: Uint8Array): Promise<boolean> {
    const prevFileContent = await this.cache.getCachedPrevFile(filePath);

    return !XmlFormatter.areEqual(newContent, prevFileContent);
  }
}
