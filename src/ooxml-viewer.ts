import { find } from 'find-in-files';
import JSZip from 'jszip';
import { lookup } from 'mime-types';
import { basename, sep } from 'path';
import vkBeautify from 'vkbeautify';
import {
  commands,
  Disposable,
  ExtensionContext,
  FileSystemWatcher,
  ProgressLocation,
  TextDocument,
  TreeView,
  Uri,
  window,
  workspace
} from 'vscode';
import xmlFormatter from 'xml-formatter';
import packageJson from '../package.json';
import { ExtensionUtilities } from './extension-utilities';
import { NORMAL_SUBFOLDER_NAME, OOXMLFileCache } from './ooxml-file-cache';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  treeView: TreeView<FileNode>;
  zip: JSZip;
  cache: OOXMLFileCache;

  watchers: Disposable[] = [];
  openTextEditors: { [key: string]: FileNode } = {};
  xmlFormatConfig = { indentation: '  ', collapseContent: true };
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
    this.treeView.title = packageJson.displayName;
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
      this.treeView.title = `${packageJson.displayName} | ${basename(file.fsPath)}`;
      await window.withProgress(
        {
          location: ProgressLocation.Notification,
          title: 'OOXML Viewer',
        },
        async progress => {
          progress.report({ message: 'Unpacking OOXML Parts' });
          await this.resetOOXMLViewer();

          // load ooxml file and populate the viewer
          const data = await this.cache.readFile(file.fsPath);
          await this.zip.loadAsync(data);
          await this.populateOOXMLViewer(this.zip.files, false);

          // set up watchers
          const watcher: FileSystemWatcher = workspace.createFileSystemWatcher(file.fsPath);
          watcher.onDidChange((uri: Uri) => {
            this.reloadOoxmlFile(file.fsPath);
          });

          const textDocumentWatcher = workspace.onDidSaveTextDocument(this.updateOOXMLFile.bind(this));

          // TODO: find a better way to remove closed text editors from the openTextEditors. The onDidCloseTextDocument takes more than 3+ minutes to fire.
          const closeWatcher = workspace.onDidCloseTextDocument((textDocument: TextDocument) => {
            delete this.openTextEditors[textDocument.fileName];
          });
          this.watchers.push(watcher, textDocumentWatcher, closeWatcher);
        },
      );
    } catch (err) {
      console.error(err);
      await window.showErrorMessage(`Could not load ${file.fsPath}`, err);
    }
  }

  /**
   * @description Displays the selected file
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
          title: 'OOXML Viewer',
        },
        async progress => {
          progress.report({ message: 'Formatting XML' });

          await this.tryFormatXml(fileNode.fullPath);

          const filePath = this.cache.getFileCachePath(fileNode.fullPath);
          this.openTextEditors[filePath] = fileNode;
          await commands.executeCommand('vscode.open', Uri.file(filePath));
        },
      );
    } catch (e) {
      console.error(e);
      await window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
    }
  }

  /**
   * @description Clears the OOXML viewer
   * @method clear
   * @returns {Promise<void>} Promise that returns void
   */
  clear(): Promise<void> {
    this.treeView.title = packageJson.displayName;
    return this.resetOOXMLViewer();
  }

  /**
   * @method getDiff
   * @async
   * @param  {FileNode} file the FileNode of the file to be diffed
   * @returns {Promise<void>}
   * @description Opens tab showing the difference between the just the primary xml part and the compare xml part
   */
  async getDiff(file: FileNode): Promise<void> {
    try {
      // format the file and its compare
      const fileContents = this.textDecoder.decode(await this.cache.getCachedFile(file.fullPath));
      if (fileContents.startsWith('<?xml')) {
        await this.cache.updateCachedFile(file.fullPath, this.textEncoder.encode(xmlFormatter(fileContents, this.xmlFormatConfig)), false);
      }

      const compareFileContents = this.textDecoder.decode(await this.cache.getCachedCompareFile(file.fullPath));
      if (compareFileContents.startsWith('<?xml')) {
        await this.cache.updateCompareFile(file.fullPath, this.textEncoder.encode(xmlFormatter(compareFileContents, this.xmlFormatConfig)));
      }

      // diff the primary and compare files
      const fileCachePath = this.cache.getFileCachePath(file.fullPath);
      const fileCompareCachePath = this.cache.getCompareFileCachePath(file.fullPath);
      const title = `${basename(fileCachePath)} ↔ compare.${basename(fileCompareCachePath)}`;

      await commands.executeCommand('vscode.diff', Uri.file(fileCompareCachePath), Uri.file(fileCachePath), title);
    } catch (err) {
      console.error(err.message || err);
      await window.showErrorMessage(err.message || err);
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
   * @description Closes all active editor tabs
   * @method closeEditors
   * @private
   * @async
   * @returns {Promise<void>}
   */
  private async closeEditors(): Promise<void> {
    return ExtensionUtilities.closeEditors(workspace.textDocuments.filter(t => t.fileName.startsWith(this.cache.cacheBasePath)));
  }

  /**
   * @description Closes all active editor tabs on VS Code opening
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
      console.error(err);
      await window.showErrorMessage('Could not remove ooxml file viewer cache');
    }
  }

  /**
   * @description Create or update tree view File Nodes and create cache files for comparison
   * @method populateOOXMLViewer
   * @private
   * @async
   * @param  {{[key:string]:JSZip.JSZipObject}} files A dictionary of files in the ooxml package.
   * @param  {boolean} showNewFileLabel Whether or not to show the new file label (asterisk) when creating a new file node.
   * @returns {Promise<void>}
   */
  private async populateOOXMLViewer(files: { [key: string]: JSZip.JSZipObject }, showNewFileLabel: boolean): Promise<void> {
    const fileKeys: string[] = Object.keys(files);
    for (const fileWithPath of fileKeys) {
      // ignore folder files
      if (files[fileWithPath].dir) {
        continue;
      }

      // Build nodes for each file
      let currentFileNode = this.treeDataProvider.rootFileNode;
      const names: string[] = fileWithPath.split('/');
      let existingFileNode: FileNode | undefined;

      for (const fileOrFolderName of names) {
        existingFileNode = currentFileNode.children.find(c => c.description === fileOrFolderName);
        if (existingFileNode) {
          currentFileNode = existingFileNode;
        } else {
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
      if (existingFileNode && !currentFileNode.isDeleted()) {
        const filesAreDifferent = await this.hasFileBeenChangedFromOutside(currentFileNode.fullPath, data);
        await this.cache.updateCachedFile(currentFileNode.fullPath, data, true);

        if (filesAreDifferent) {
          currentFileNode.setModified();
        } else {
          currentFileNode.setUnchanged();
        }
      } else {
        await this.cache.createCachedFile(currentFileNode.fullPath, data, showNewFileLabel);
        if (showNewFileLabel) {
          currentFileNode.setCreated();
        }
      }
    }

    await this.removeDeletedParts();
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
   * @param  {TextDocument} document The text document.
   * @returns {Promise<void>}
   */
  private async updateOOXMLFile(document: TextDocument): Promise<void> {
    try {
      const cacheFilePath = document.fileName;
      if (!this.cache.pathBelongsToCache(cacheFilePath) || !this.cache.cachePathIsNormal(cacheFilePath)) {
        return;
      }

      const filePath = this.cache.getFilePathFromCacheFilePath(cacheFilePath);

      const fileContents = await this.cache.getCachedFile(filePath);
      const prevFileContents = await this.cache.getCachedPrevFile(filePath);
      const fileMinXml = vkBeautify.xmlmin(this.textDecoder.decode(fileContents), true);
      const prevFileMinXml = vkBeautify.xmlmin(this.textDecoder.decode(prevFileContents), true);

      if (Buffer.from(fileMinXml).equals(Buffer.from(prevFileMinXml))) {
        return;
      }
      const mimeType = lookup(basename(this.ooxmlFilePath)) || undefined;
      const zipFile = await this.zip
        .file(filePath, this.textEncoder.encode(fileMinXml))
        .generateAsync({ type: 'uint8array', mimeType, compression: 'DEFLATE' });
      await this.cache.writeFile(this.ooxmlFilePath, zipFile, true);

      await this.cache.createCachedFile(filePath, fileContents, false);
      this.treeDataProvider.refresh();
    } catch (err) {
      if (err?.code === 'EBUSY' || err?.message.toLowerCase().includes('ebusy')) {
        await window.showWarningMessage(
          `File not saved.\n${basename(this.ooxmlFilePath)} is open in another program.\nClose that program before making any changes.`,
          { modal: true },
        );

        await ExtensionUtilities.makeTextEditorDirty(window.activeTextEditor);
      } else {
        console.error(err.message || err);
        await window.showErrorMessage(err.message || err);
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
        title: 'OOXML Viewer',
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
          console.error(err.message);
          await window.showErrorMessage(err.message || err);
        }
      },
    );
  }

  /**
   * Traverse tree and delete cached parts that don't exist
   *
   * @description Delete cache files for parts deleted from OOXML file
   * @method removeDeletedParts
   * @async
   * @private
   * @returns {Promise<void>}
   */
  private async removeDeletedParts(): Promise<void> {
    try {
      const filesInOoxmlFile = new Set(Object.keys(this.zip.files));
      const fileNodeQueue = [this.treeDataProvider.rootFileNode];

      let fileNode;
      while ((fileNode = fileNodeQueue.pop())) {
        if (fileNode.fullPath && !filesInOoxmlFile.has(fileNode.fullPath)) {
          const cachedFileExists = await this.cache.getCachedFile(fileNode.fullPath);
          if (cachedFileExists && !fileNode.isDeleted()) {
            fileNode.setDeleted();
            await this.cache.updateCachedFile(fileNode.fullPath, new Uint8Array(), true);
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
      console.error(err.message || err);
      await window.showErrorMessage(err.message || err);
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
            await this.tryFormatXml(filePath);
          } catch (err) {
            console.error(err);
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
      console.error(err.message || err);
      await window.showErrorMessage(err.message || err);
    }
  }

  /**
   * @description Tries to format the file as xml.
   * @method tryFormatXml
   * @async
   * @private
   * @param {string} filePath The path of the file in the ooxml package.
   * @returns {Promise<boolean>} A Promise resolving to whether or not the file has been formatted.
   */
  private async tryFormatXml(filePath: string) {
    const xml = this.textDecoder.decode(await this.cache.getCachedFile(filePath));
    if (xml.startsWith('<?xml')) {
      const text = xmlFormatter(xml, this.xmlFormatConfig);
      await this.cache.updateCachedFile(filePath, this.textEncoder.encode(text), false);
      return true;
    }

    return false;
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
      const fileArrayBuffer = newContent;
      const prevFileArrayBuffer = await this.cache.getCachedPrevFile(filePath);
      if (Buffer.from(fileArrayBuffer).equals(Buffer.from(prevFileArrayBuffer))) {
        return false;
      }

      const fileText = this.textDecoder.decode(fileArrayBuffer);
      const prevFileText = this.textDecoder.decode(prevFileArrayBuffer);
      const minFileText = vkBeautify.xmlmin(fileText);
      const minPrevFileText = vkBeautify.xmlmin(prevFileText);
      return minFileText !== minPrevFileText;
    } catch (err) {
      console.error(err.message || err);
      await window.showErrorMessage(err.message || err);
    }

    return false;
  }

  /**
   * @description search OOXML parts for a string and display the results in a web view
   * @method searchOoxmlParts
   * @returns {Promise<void>}
   */
  async searchOoxmlParts(): Promise<void> {
    const warningMsg = 'A file must be open in the OOXML Viewer to search its parts.';
    try {
      await workspace.fs.stat(Uri.file(this.cache.normalSubfolderPath));
      const searchTerm = await window.showInputBox({ title: 'Search OOXML Parts', prompt: 'Enter a search term.' });
      if (!searchTerm) {
        return;
      }
      const results = await find(searchTerm, this.cache.normalSubfolderPath);

      for (const filePath in results) {
        const ooxmlPath = filePath.split(NORMAL_SUBFOLDER_NAME)[1].split(sep).join('/');
        await this.tryFormatXml(ooxmlPath);
      }

      await commands.executeCommand('workbench.action.findInFiles', {
        query: searchTerm,
        filesToInclude: this.cache.normalSubfolderPath,
        triggerSearch: true,
        isCaseSensitive: false,
        matchWholeWord: false,
      });
    } catch (err) {
      if (err?.code === 'ENOENT' || err?.code === 'FileNotFound') {
        window.showWarningMessage(warningMsg);
      } else {
        const msg = err.message || err;
        console.error(msg);
        window.showErrorMessage(msg);
      }
    }
  }
}
