import JSZip from 'jszip';
import { basename } from 'path';
import vkBeautify from 'vkbeautify';
import {
  commands,
  Disposable,
  ExtensionContext,
  FileSystemWatcher,
  ProgressLocation,
  TextDocument,
  Uri,
  window,
  workspace
} from 'vscode';
import { ExtensionUtilities } from './extension-utilities';
import { OOXMLFileCache } from './ooxml-file-cache';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  zip: JSZip;
  watchers: Disposable[] = [];
  openTextEditors: { [key: string]: FileNode } = {};
  ooxmlFilePath = '';
  cache: OOXMLFileCache;
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
    this.zip = new JSZip();
    this.cache = new OOXMLFileCache(context);
    this.closeEditors();
  }

  /**
   * @description Loads the selected OOXML file into the tree view and add file listeners
   * @method viewContents
   * @async
   * @param {Uri} file The OOXML file
   * @returns {Promise<void>}
   */
  async viewContents(file: Uri): Promise<void> {
    try {
      this.ooxmlFilePath = file.fsPath;
      await window.withProgress(
        {
          location: ProgressLocation.Notification,
          title: 'OOXML Viewer',
        },
        async progress => {
          progress.report({ message: 'Unpacking OOXML Parts' });
          await this.resetOOXMLViewer();
          const data = await workspace.fs.readFile(Uri.file(file.fsPath));
          await this.zip.loadAsync(data);
          await this.populateOOXMLViewer(this.zip.files, false);

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
      window.showErrorMessage(`Could not load ${file.fsPath}`, err);
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

          this.tryFormatXml(fileNode.fullPath);

          const filePath = this.cache.getFileCachePath(fileNode.fullPath);
          this.openTextEditors[filePath] = fileNode;
          commands.executeCommand('vscode.open', Uri.file(filePath));
        },
      );
    } catch (e) {
      console.error(e);
      window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
    }
  }

  /**
   * @description Clears the OOXML viewer
   * @method clear
   * @returns {Promise<void>} Promise that returns void
   */
  clear(): Promise<void> {
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
      // get the full path for the primary file and the compare files
      const fileCachePath = this.cache.getFileCachePath(file.fullPath);
      const fileCompareCachePath = this.cache.getCompareFileCachePath(file.fullPath);
      const fileContents = this.textDecoder.decode(await this.cache.getCachedFile(file.fullPath));
      const compareFileContents = this.textDecoder.decode(await this.cache.getCachedCompareFile(file.fullPath));

      await this.cache.updateCachedFile(file.fullPath, this.textEncoder.encode(fileContents.startsWith('<?xml')
        ? vkBeautify.xml(fileContents)
        : fileContents), false);
      await this.cache.updateCompareFile(file.fullPath, this.textEncoder.encode(compareFileContents.startsWith('<?xml')
        ? vkBeautify.xml(compareFileContents)
        : compareFileContents));

      // diff the primary and compare files
      const title = `${basename(fileCachePath)} ↔ ${basename(fileCompareCachePath)}`;
      await commands.executeCommand('vscode.diff', Uri.file(fileCompareCachePath), Uri.file(fileCachePath), title);
    } catch (err) {
      console.error(err);
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
    return ExtensionUtilities.closeEditors(
      workspace.textDocuments.filter(t => t.fileName.startsWith(this.cache.cacheBasePath)));
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

      await Promise.all([
        this.closeEditors(),
        this.cache.clear(),
      ]);
    } catch (err) {
      console.error(err);
      window.showErrorMessage('Could not remove ooxml file viewer cache');
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
      const data = await this.zip.file(currentFileNode.fullPath)?.async('uint8array') ?? new Uint8Array();
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
      if (!this.cache.pathBelongsToCache(cacheFilePath) || !this.cache.pathIsNotPrevOrCompare(cacheFilePath)) {
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
      
      const zipFile = await this.zip.file(filePath, fileMinXml, { binary: true }).generateAsync({ type: 'uint8array' });
      await workspace.fs.writeFile(Uri.file(this.ooxmlFilePath), zipFile);

      await this.cache.createCachedFile(filePath, fileContents, false);
      this.treeDataProvider.refresh();
    } catch (err) {
      if (err?.code === 'EBUSY' || err?.message.toLowerCase().includes('ebusy')) {
        window.showWarningMessage(
          `File not saved.\n${basename(this.ooxmlFilePath)} is open in another program.\nClose that program before making any changes.`,
          { modal: true },
        );

        await ExtensionUtilities.makeTextEditorDirty(window.activeTextEditor);
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
          const data = await workspace.fs.readFile(Uri.file(filePath));
          this.zip = new JSZip();
          await this.zip.loadAsync(data);

          await this.populateOOXMLViewer(this.zip.files, true);
        } catch (err) {
          console.error(err);
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
      console.error(err);
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
      workspace.textDocuments
        .filter(d => this.cache.pathBelongsToCache(d.fileName))
        .filter(d => Object.keys(this.zip.files).filter(f => f.includes(basename(d.fileName))))
        .forEach(async d => {
          try {
            const filePath = this.cache.getFilePathFromCacheFilePath(d.fileName);
            await this.tryFormatXml(filePath);
          } catch (err) {
            console.error(err);
          }
        });

      workspace.textDocuments
        .filter(d => this.cache.pathBelongsToCache(d.fileName))
        .filter(d => !Object.keys(this.zip.files).filter(f => f.includes(basename(d.fileName))))
        .forEach(async d => {
          await window.showTextDocument(Uri.file(d.fileName), { preview: true, preserveFocus: false });
          await commands.executeCommand('workbench.action.closeActiveEditor');
        });
    } catch (err) {
      console.error(err);
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
      const text = vkBeautify.xml(xml);
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
    }

    return false;
  }
}
