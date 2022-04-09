import JSZip from 'jszip';
import { lookup } from 'mime-types';
import { basename, dirname } from 'path';
import vkBeautify from 'vkbeautify';
import {
  commands,
  Disposable,
  ExtensionContext,
  FileSystemError,
  FileSystemWatcher,
  ProgressLocation,
  RelativePattern,
  TextDocument,
  TreeView,
  Uri,
  window,
  workspace,
} from 'vscode';
import xmlFormatter from 'xml-formatter';
import packageJson from '../package.json';
import { ExtensionUtilities } from './extension-utilities';
import { OOXMLFileCache } from './ooxml-file-cache';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';

const extensionName = packageJson.displayName;

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  cache: OOXMLFileCache;
  zip: JSZip;
  treeView: TreeView<FileNode>;

  watchers: Disposable[] = [];
  ooxmlFilePath = '';

  textEncoder = new TextEncoder();
  textDecoder = new TextDecoder();

  /**
   * @description Constructs an instance of OOXMLViewer
   * @constructor OOXMLViewer
   * @param  {ExtensionContext} context The extension context
   * @returns {OOXMLViewer} instance
   */
  constructor(context: ExtensionContext) {
    this.treeDataProvider = new OOXMLTreeDataProvider();
    this.treeView = window.createTreeView('ooxmlViewer', { treeDataProvider: this.treeDataProvider });
    this.treeView.title = extensionName;
    context.subscriptions.push(this.treeView);
    this.zip = new JSZip();
    this.cache = new OOXMLFileCache(context);

    this.closeEditorsOnStartup();
  }

  /**
   * @description Loads the selected OOXML file into the tree view and add file listeners
   * @method openOoxmlPackage
   * @async
   * @param {Uri} file The OOXML file
   * @returns {Promise<void>}
   */
  async openOoxmlPackage(file: Uri): Promise<void> {
    try {
      this.ooxmlFilePath = file.fsPath;
      this.treeView.title = `${extensionName} | ${basename(file.fsPath)}`;
      await window.withProgress(
        {
          location: ProgressLocation.Notification,
          title: extensionName,
        },
        async progress => {
          progress.report({ message: 'Unpacking OOXML Parts' });
          await this.resetOOXMLViewer();

          // load ooxml file and populate the viewer
          const data = await this.cache.readFile(file.fsPath);
          await this.zip.loadAsync(data);
          await this.populateOOXMLViewer(this.zip.files, false);

          // set up watchers
          const fileSystemWatcher: FileSystemWatcher = workspace.createFileSystemWatcher(
            new RelativePattern(dirname(file.fsPath), basename(file.fsPath)),
          );

          // Prevent multiple comparison operations on large files
          let locked = false;
          fileSystemWatcher.onDidChange(async (_: Uri) => {
            if (!locked) {
              locked = true;
              await this.reloadOoxmlFile(file.fsPath);
              locked = false;
            }
          });

          const openTextDocumentWatcher = workspace.onDidOpenTextDocument(document => this.tryFormatDocument(document.fileName));
          const saveTextDocumentWatcher = workspace.onDidSaveTextDocument(document => this.updateOOXMLFile(document));

          this.watchers.push(openTextDocumentWatcher, fileSystemWatcher, saveTextDocumentWatcher);
        },
      );
    } catch (err) {
      await ExtensionUtilities.handleError(err);
    }
  }

  /**
   * @description Displays and formats the selected file
   * @async
   * @method viewFile
   * @param {FileNode} fileNode The selected file node
   * @returns {Promise<void>}
   */
  async viewFile(fileNode: FileNode): Promise<void> {
    try {
      await window.withProgress(
        {
          location: ProgressLocation.Notification,
          title: extensionName,
        },
        async progress => {
          progress.report({ message: 'Formatting XML' });

          await this.formatXml(fileNode.fullPath);

          const filePath = this.cache.getNormalFileCachePath(fileNode.fullPath);
          await commands.executeCommand('vscode.open', Uri.file(filePath));
        },
      );
    } catch (err) {
      await ExtensionUtilities.handleError(err);
    }
  }

  /**
   * @description Clears the OOXML viewer
   * @method clear
   * @returns {Promise<void>} Promise that returns void
   */
  clear(): Promise<void> {
    this.treeView.title = extensionName;
    return this.resetOOXMLViewer();
  }

  /**
   * @description Opens tab showing the difference between the primary xml part and the compare xml part
   * @method getDiff
   * @async
   * @param  {FileNode} file the FileNode of the file to be diffed
   * @returns {Promise<void>}
   */
  async getDiff(file: FileNode): Promise<void> {
    try {
      // format the file
      await this.formatXml(file.fullPath);

      // diff the primary and compare files
      const fileCachePath = this.cache.getNormalFileCachePath(file.fullPath);
      const fileCompareCachePath = this.cache.getCompareFileCachePath(file.fullPath);
      const title = `${basename(fileCachePath)} ↔ compare.${basename(fileCompareCachePath)}`;

      await commands.executeCommand('vscode.diff', Uri.file(fileCompareCachePath), Uri.file(fileCachePath), title);
    } catch (err) {
      await ExtensionUtilities.handleError(err);
    }
  }

  /**
   * @description Format the document if it is a cached normal file
   * @method tryFormatDocument
   * @async
   * @returns Promise{void}
   */
  async tryFormatDocument(filePath: string): Promise<void> {
    if (this.cache.cachePathIsNormal(filePath)) {
      await window.withProgress(
        {
          location: ProgressLocation.Notification,
          title: extensionName,
        },
        async progress => {
          progress.report({ message: 'Formatting XML' });
          await this.formatXml(this.cache.getFilePathFromCacheFilePath(filePath));
        },
      );
    }
  }

  /**
   * @description Search OOXML parts for a string and display the results in a web view
   * @method searchOoxmlParts
   * @returns {Promise<void>}
   */
  async searchOoxmlParts(): Promise<void> {
    const warningMsg = `A file must be open in the ${extensionName} to search its parts.`;
    try {
      await workspace.fs.stat(Uri.file(this.cache.normalSubfolderPath));
      const searchTerm = await window.showInputBox({ title: 'Search OOXML Parts', prompt: 'Enter a search term.' });

      if (!searchTerm) {
        return;
      }

      await commands.executeCommand('workbench.action.findInFiles', {
        query: searchTerm,
        filesToInclude: this.cache.normalSubfolderPath,
        triggerSearch: true,
        isCaseSensitive: false,
        matchWholeWord: false,
      });
    } catch (err) {
      if ((err as FileSystemError)?.code?.toLowerCase() === 'filenotfound') {
        window.showWarningMessage(warningMsg);
      } else {
        await ExtensionUtilities.handleError(err);
      }
    }
  }

  /**
   * @description Close all file watchers
   * @method disposeWatchers
   * @async
   * @returns {void}
   */
  disposeWatchers(): void {
    if (this.watchers.length) {
      this.watchers.forEach(w => w.dispose());
      this.watchers = [];
    }
  }

  /**
   * @description Closes all active editor tabs that contain xml parts
   * @method closeEditors
   * @private
   * @async
   * @returns {Promise<void>}
   */
  private async closeEditors(): Promise<void> {
    return ExtensionUtilities.closeEditors(workspace.textDocuments.filter(t => t.fileName.startsWith(this.cache.cacheBasePath)));
  }

  /**
   * @description Closes all active editor tabs that contain xml parts on VS Code opening
   * @method closeEditors
   * @private
   * @async
   * @returns {Promise<void>}
   */
  private async closeEditorsOnStartup(): Promise<void> {
    return ExtensionUtilities.closeEditorsOnStartup(this.cache.cacheBasePath);
  }

  /**
   * @description Sets this.zip to an empty zip file, deletes the cache folder, closes all watchers, and closes all editor tabs
   * @method resetOOXMLViewer
   * @private
   * @async
   * @returns {Promise<void>}
   */
  private async resetOOXMLViewer(): Promise<void> {
    try {
      this.zip = new JSZip();
      this.treeDataProvider.rootFileNode = new FileNode();
      this.treeDataProvider.refresh();
      this.disposeWatchers();

      await Promise.all([this.closeEditors(), this.cache.clear()]);
    } catch (err) {
      await ExtensionUtilities.handleError(err);
    }
  }

  /**
   * @description Create or update tree view File Nodes and create cache files for comparison
   * @method populateOOXMLViewer
   * @private
   * @async
   * @param  {{[key:string]:JSZip.JSZipObject}} files A dictionary of files in the ooxml package.
   * @param  {boolean} isFirstOpen Whether or not this is the first time the file has been populated
   *  (Should the new file label (asterisk) be shown when creating a new file node).
   * @returns {Promise<void>}
   */
  private async populateOOXMLViewer(files: { [key: string]: JSZip.JSZipObject }, isFirstOpen: boolean): Promise<void> {
    const fileKeys: string[] = Object.keys(files);
    for (const fileWithPath of fileKeys) {
      // ignore folder files
      if (files[fileWithPath].dir) {
        continue;
      }

      // Build nodes for each file in the package

      let fileNodeAlreadyExists = true;
      let currentFileNode = this.treeDataProvider.rootFileNode;
      const names: string[] = fileWithPath.split('/');
      for (const fileOrFolderName of names) {
        const existingFileNode = currentFileNode.children.find(c => c.description === fileOrFolderName);
        if (existingFileNode) {
          currentFileNode = existingFileNode;
        } else {
          fileNodeAlreadyExists = false;
          // create a new FileNode with the currentFileNode as parent and add it to the currentFileNode children
          const newFileNode = new FileNode();
          newFileNode.fileName = fileOrFolderName;
          newFileNode.parent = currentFileNode;
          currentFileNode.children.push(newFileNode);
          currentFileNode = newFileNode;
        }
      }

      // set the current node fullPath (either new or existing) to fileWithPath
      currentFileNode.fullPath = fileWithPath;

      // cache or update the cache of the node and mark the status of the node
      const data = (await this.zip.file(currentFileNode.fullPath)?.async('uint8array')) ?? new Uint8Array();

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
        const filesAreDifferent = await this.hasFileBeenChangedFromOutside(currentFileNode.fullPath, data);
        await this.cache.updateCachedFiles(currentFileNode.fullPath, data);

        if (filesAreDifferent) {
          currentFileNode.setModified();
        } else {
          currentFileNode.setUnchanged();
        }
      } else {
        if (isFirstOpen) {
          await this.cache.createCachedFilesWithEmptyCompare(currentFileNode.fullPath, data);
          currentFileNode.setCreated();
        } else {
          await this.cache.createCachedFiles(currentFileNode.fullPath, data);
        }
      }
    }

    // need to handle deleted parts separately since the zip
    // doesn't contain them anymore
    await this.handleDeletedParts();
    await this.reformatOpenTabs();

    // tell vscode the tree has changed
    this.treeDataProvider.refresh();
  }

  /**
   * @description Writes changes to OOXML file being inspected when one of its parts is saved.
   * Note that this will trigger `reloadOoxmlFile` to fire if changes are written.
   * @method updateOOXMLFile
   * @async
   * @private
   * @param  {TextDocument} document The text document of the xml part that was updated.
   * @returns {Promise<void>}
   */
  private async updateOOXMLFile(document: TextDocument): Promise<void> {
    try {
      const cacheFilePath = document.fileName;
      if (!this.cache.pathBelongsToCache(cacheFilePath) || !this.cache.cachePathIsNormal(cacheFilePath)) {
        return;
      }

      const filePath = this.cache.getFilePathFromCacheFilePath(cacheFilePath);

      const fileContents = await this.cache.getCachedNormalFile(filePath);
      const prevFileContents = await this.cache.getCachedPrevFile(filePath);
      if (this.isXmlEqual(fileContents, prevFileContents)) {
        return;
      }

      const fileMinXml = vkBeautify.xmlmin(this.textDecoder.decode(fileContents), true);
      const mimeType = lookup(basename(this.ooxmlFilePath)) || undefined;
      const zipFile = await this.zip
        .file(filePath, this.textEncoder.encode(fileMinXml))
        .generateAsync({ type: 'uint8array', mimeType, compression: 'DEFLATE' });
      await this.cache.writeFile(this.ooxmlFilePath, zipFile, true);
      await this.cache.createCachedFiles(filePath, fileContents);

      this.treeDataProvider.refresh();
    } catch (err) {
      if ((err as FileSystemError)?.code.toLowerCase() === 'unknown' && (err as FileSystemError)?.message.toLowerCase().includes('ebusy')) {
        await window.showWarningMessage(
          `File not saved.\n${basename(this.ooxmlFilePath)} is open in another program.\nClose that program before making any changes.`,
          { modal: true },
        );

        await ExtensionUtilities.makeTextEditorDirty(window.activeTextEditor);
      } else {
        await ExtensionUtilities.handleError(err);
      }
    }
  }

  /**
   * @description Update the OOXML cache files
   * @method reloadOoxmlFile
   * @async
   * @private
   * @param  {string} filePath Path to the OOXML file to load
   * @returns {Promise<void>}
   */
  private async reloadOoxmlFile(filePath: string): Promise<void> {
    await window.withProgress(
      {
        location: ProgressLocation.Notification,
        title: extensionName,
      },
      async progress => {
        try {
          progress.report({ message: 'Updating OOXML Parts' });

          // unzip ooxml file again
          const data = await this.cache.readFile(filePath);
          this.zip = new JSZip();
          await this.zip.loadAsync(data);

          await this.populateOOXMLViewer(this.zip.files, true);
        } catch (err) {
          await ExtensionUtilities.handleError(err);
        }
      },
    );
  }

  /**
   * @description Traverse tree and delete cached parts that don't exist anymore.
   * If the file node is marked as deleted, delete the cached part, if it isn't then
   * mark the node for deletion on next update
   * @method handleDeletedParts
   * @async
   * @private
   * @returns {Promise<void>}
   */
  private async handleDeletedParts(): Promise<void> {
    try {
      const filesInOoxmlFile = new Set(Object.keys(this.zip.files));
      const fileNodeQueue = [this.treeDataProvider.rootFileNode];

      let fileNode;
      while ((fileNode = fileNodeQueue.pop())) {
        if (fileNode.fullPath && !filesInOoxmlFile.has(fileNode.fullPath)) {
          if (!fileNode.isDeleted()) {
            fileNode.setDeleted();
            await this.cache.updateCachedFiles(fileNode.fullPath, new Uint8Array());
          } else {
            // remove files marked as deleted from tree view and cache after the ooxml file
            // the second time the ooxml file is saved
            await this.cache.deleteCachedFiles(fileNode.fullPath);
            fileNode.parent?.children.splice(fileNode.parent.children.indexOf(fileNode), 1);
          }
        }

        fileNodeQueue.push(...fileNode.children);
      }

      this.treeDataProvider.refresh();
    } catch (err) {
      await ExtensionUtilities.handleError(err);
    }
  }

  /**
   * @description Reformats the open tabs after their contents have been updated
   * @method reformatOpenTabs
   * @async
   * @private
   * @returns {Promise<void>}
   */
  private async reformatOpenTabs(): Promise<void> {
    try {
      const filePathsInOoxmlPackage = new Set(Object.keys(this.zip.files));
      workspace.textDocuments
        .filter(d => this.cache.pathBelongsToCache(d.fileName))
        .map(d => this.cache.getFilePathFromCacheFilePath(d.fileName))
        .filter(p => filePathsInOoxmlPackage.has(p))
        .forEach(async filePath => {
          try {
            await this.formatXml(filePath);
          } catch (err) {
            await ExtensionUtilities.handleError(err);
          }
        });

      workspace.textDocuments
        .filter(d => this.cache.pathBelongsToCache(d.fileName))
        .filter(d => !filePathsInOoxmlPackage.has(this.cache.getFilePathFromCacheFilePath(d.fileName)))
        .forEach(async d => {
          await window.showTextDocument(Uri.file(d.fileName), { preview: true, preserveFocus: false });
          await commands.executeCommand('workbench.action.closeActiveEditor');
        });
    } catch (err) {
      await ExtensionUtilities.handleError(err);
    }
  }

  /**
   * @description Tries to format the file as xml
   * and if successful, updates the cached files.
   * @method formatXml
   * @async
   * @private
   * @param {string} filePath The path of the file in the ooxml package.
   * @returns {Promise<boolean>} A Promise resolving to whether or not the file has been formatted.
   */
  private async formatXml(filePath: string): Promise<void> {
    // Need to format the normal/prev and the compare separately for the diff to work
    const xmlFormatConfig = { indentation: '  ', collapseContent: true };
    const formatNormalXml = async () => {
      const fileContent = this.textDecoder.decode(await this.cache.getCachedNormalFile(filePath));
      if (fileContent.startsWith('<?xml')) {
        // for some reason xmlFormatter doesn't always format everything without minifying it first
        const formattedXml = xmlFormatter(vkBeautify.xmlmin(fileContent), xmlFormatConfig);
        await this.cache.updateCachedFilesNoCompare(filePath, this.textEncoder.encode(formattedXml));
      }
    };

    const formatCompareXml = async () => {
      const compareFileContent = this.textDecoder.decode(await this.cache.getCachedCompareFile(filePath));
      if (compareFileContent.startsWith('<?xml')) {
        // for some reason xmlFormatter doesn't always format everything without minifying it first
        const formattedXml = xmlFormatter(vkBeautify.xmlmin(compareFileContent), xmlFormatConfig);
        await this.cache.updateCompareFile(filePath, this.textEncoder.encode(formattedXml));
      }
    };

    await Promise.all([formatNormalXml(), formatCompareXml()]);
  }

  /**
   * @description Check if an OOXML part is different from its cached version.
   * @method hasFileBeenChangedFromOutside
   * @async
   * @private
   * @param {string} filePath The path of the file in the ooxml package.
   * @param {string} newContent The updated contents of the file.
   * @returns {Promise<boolean>} A Promise resolving to whether or not the file has been changed from the outside.
   */
  private async hasFileBeenChangedFromOutside(filePath: string, newContent: Uint8Array): Promise<boolean> {
    try {
      // fast check if contents are exactly equals
      const prevFileContent = await this.cache.getCachedPrevFile(filePath);
      if (Buffer.from(newContent).equals(Buffer.from(prevFileContent))) {
        return false;
      }

      return !this.isXmlEqual(newContent, prevFileContent);
    } catch (err) {
      await ExtensionUtilities.handleError(err);
    }

    return false;
  }

  /**
   * @description Check if two xml files are the same.
   * @method areXmlEqual
   * @async
   * @private
   * @param {string} xmlContent1 the first xml content to compare.
   * @param {string} xmlContent2 The second xml content to compare.
   * @returns {Promise<boolean>} A Promise resolving to whether or not the xml contents are the same.
   */
  private isXmlEqual(xmlContent1: Uint8Array, xmlContent2: Uint8Array): boolean {
    const fileMinXml = vkBeautify.xmlmin(this.textDecoder.decode(xmlContent1), true);
    const prevFileMinXml = vkBeautify.xmlmin(this.textDecoder.decode(xmlContent2), true);

    if (fileMinXml === prevFileMinXml) {
      return true;
    }

    return false;
  }
}
