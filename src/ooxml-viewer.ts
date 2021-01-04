import { existsSync } from 'fs';
import JSZip from 'jszip';
import { basename } from 'path';
import vkBeautify from 'vkbeautify';
import {
  commands,
  Disposable,
  ExtensionContext,
  FileStat,
  FileSystemWatcher,
  FileType,
  ProgressLocation,
  TextDocument,
  Uri,
  window,
  workspace
} from 'vscode';
import { ExtensionUtilities } from './extension-utilities';
import { OOXMLFileCache } from './ooxml-file-cache';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';

const MAXIMUM_XML_PARSING_CHARACTERS = 1000000;

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  zip: JSZip;
  watchers: Disposable[] = [];
  watchActions: { [key: string]: number } = {};
  openTextEditors: { [key: string]: FileNode } = {};
  ooxmlFilePath = '';
  cache: OOXMLFileCache;

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
          await this._populateOOXMLViewer(this.zip.files, false);

          const watcher: FileSystemWatcher = workspace.createFileSystemWatcher(file.fsPath);

          watcher.onDidChange((uri: Uri) => {
            this._reloadOoxmlFile(file.fsPath);
          });

          const textDocumentWatcher = workspace.onDidSaveTextDocument(this._updateOOXMLFile.bind(this));

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

          await this._tryFormatXMLFile(fileNode.fullPath);

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
      const fileContents = (await this.cache.getFile(file.fullPath)).toString();
      const compareFileContents = (await this.cache.getCompareFile(file.fullPath)).toString();

      const enc = new TextEncoder();
      await this.cache.cacheFile(file.fullPath, enc.encode(fileContents.startsWith('<?xml') ? vkBeautify.xml(fileContents) : fileContents));
      await this.cache.cacheCompareFile(file.fullPath, enc.encode(compareFileContents.startsWith('<?xml')
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
   * @param  {TextDocument[]} textDocuments
   * @returns {Promise<void>}
   */
  async closeEditors(): Promise<void> {
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
   * @method _populateOOXMLViewer
   * @private
   * @async
   * @param  {{[key:string]:JSZip.JSZipObject}} files
   * @param  {boolean} showNewFileLabel
   * @returns {Promise<void>}
   */
  private async _populateOOXMLViewer(files: { [key: string]: JSZip.JSZipObject }, showNewFileLabel: boolean): Promise<void> {
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
        // Create node if it does not exist
        existingFileNode = currentFileNode.children.find(c => c.description === fileOrFolderName);
        if (existingFileNode) {
          currentFileNode = existingFileNode;
        } else {
          // create a new FileNode with the currentFileNode as parent and add it to the currentFileNode children
          const newFileNode = new FileNode();
          newFileNode.fileName = fileOrFolderName;
          newFileNode.parent = currentFileNode;
          // newFileNode.fullPath = fileWithPath;
          currentFileNode.children.push(newFileNode);
          currentFileNode = newFileNode;
        }
      }

      // set the current node fullPath (either new or existing) to fileWithPath
      currentFileNode.fullPath = fileWithPath;

      // cache or update the cache of the node and mark the status of the node
      const data = await this.zip.file(currentFileNode.fullPath)?.async('uint8array') ?? new Uint8Array();
      if (existingFileNode) {
        const filesAreDifferent = await this._hasFileBeenChangedFromOutside(currentFileNode.fullPath, data);
        await this.cache.updateCachedFile(currentFileNode.fullPath, data);

        if (filesAreDifferent) {
          currentFileNode.setModified();
        } else {
          currentFileNode.setUnchanged();
        }
      } else {
        await this.cache.cacheCachedFile(currentFileNode.fullPath, data, showNewFileLabel);
        if (showNewFileLabel) {
          currentFileNode.setCreated();
        }
      }
    }

    await this._removeDeletedParts();
    await this._reformatOpenTabs();

    // tell vscode the tree has changed
    this.treeDataProvider.refresh();
  }

  private async _tryFormatXMLFile(filePath: string): Promise<boolean> {
    const data = await this.cache.getFile(filePath);
    const text = new TextDecoder().decode(data);
    if (!text.startsWith('<?xml')) {
      return false;
    }

    const lengthWithNoWhitespace = text.replace(/\s+/g, '').length;
    if (lengthWithNoWhitespace >= MAXIMUM_XML_PARSING_CHARACTERS) {
      window.showWarningMessage(
        `${basename(filePath)} is too large to format.\nOOXML Parts must be less than ` +
        `${MAXIMUM_XML_PARSING_CHARACTERS.toLocaleString()} characters to format`,
        { modal: true },
      );

      return false;
    }

    const formattedXml = vkBeautify.xml(text);
    await this.cache.cacheFile(filePath, new TextEncoder().encode(formattedXml || text));

    return true;
  }

  /**
   * @description Writes changes to OOXML file being inspected
   * @method _updateOOXMLFile
   * @async
   * @private
   * @param  {TextDocument} document The text document
   * @returns {Promise<void>}
   */
  private async _updateOOXMLFile(document: TextDocument): Promise<void> {
    try {
      const filePath = document.fileName;
      const prevFilePath = this.cache.getPrevFilePath(filePath);

      if (!(filePath && existsSync(filePath) && prevFilePath && existsSync(prevFilePath))) {
        return;
      }

      const stats: FileStat = await workspace.fs.stat(Uri.file(filePath));
      const time = stats.mtime;

      if (stats.type === FileType.Directory || this.watchActions[filePath] === time) {
        return;
      }

      this.watchActions[filePath] = time;

      const textDecoder = new TextDecoder();
      const fileContents = await workspace.fs.readFile(Uri.file(filePath));
      const prevFileContents = await workspace.fs.readFile(Uri.file(prevFilePath));
      const fileMinXml = vkBeautify.xmlmin(textDecoder.decode(fileContents), true);
      const prevFileMinXml = vkBeautify.xmlmin(textDecoder.decode(prevFileContents), true);

      if (Buffer.from(fileMinXml).equals(Buffer.from(prevFileMinXml))) {
        return;
      }

      let normalizedPath = filePath.substring(this.cache.cacheBasePath.length).replace(/\\/g, '/');
      normalizedPath = normalizedPath.startsWith('/') ? normalizedPath.substring(1) : normalizedPath;
      const zipFile = await this.zip.file(normalizedPath, fileMinXml, { binary: true }).generateAsync({ type: 'uint8array' });
      await workspace.fs.writeFile(Uri.file(this.ooxmlFilePath), zipFile);

      await workspace.fs.writeFile(Uri.file(this.cache.getCompareFilePath(filePath)), prevFileContents);
      await workspace.fs.writeFile(Uri.file(this.cache.getPrevFilePath(filePath)), fileContents);
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
   * @method _reloadOoxmlFile
   * @async
   * @private
   * @param  {string} filePath Path to the OOXML file to load
   * @returns {Promise<void>}
   */
  private async _reloadOoxmlFile(filePath: string): Promise<void> {
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

          await this._populateOOXMLViewer(this.zip.files, true);
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
   * @method _removeDeletedParts
   * @async
   * @private
   * @param  {FileNode} node?
   * @returns {Promise<void>}
   */
  private async _removeDeletedParts(node?: FileNode): Promise<void> {
    try {
      const filesInOoxmlFile = new Set(Object.keys(this.zip.files));
      const fileNodeQueue = [this.treeDataProvider.rootFileNode];

      let fileNode;
      while ((fileNode = fileNodeQueue.pop())) {
        if (fileNode.fullPath && !filesInOoxmlFile.has(fileNode.fullPath)) {
          const cachedFileExists = await this.cache.readFile(fileNode.fullPath);
          if (cachedFileExists) {
            fileNode.setDeleted();
            await this.cache.cacheFile(fileNode.fullPath, new Uint8Array());
          } else {
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
   * @method _reformatOpenTabs
   * @async
   * @private
   * @returns {Promise<void>}
   */
  private async _reformatOpenTabs(): Promise<void> {
    try {
      workspace.textDocuments
        .filter(d => d.fileName.startsWith(this.cache.cacheBasePath))
        .filter(d => Object.keys(this.zip.files).filter(f => f.includes(basename(d.fileName))))
        .forEach(async d => {
          try {
            const xml = (await workspace.fs.readFile(Uri.file(d.fileName))).toString();
            if (xml.startsWith('<?xml')) {
              const text = vkBeautify.xml(xml);
              await workspace.fs.writeFile(Uri.file(d.fileName), new TextEncoder().encode(text));
            }
          } catch (err) {
            console.error(err);
          }
        });

      workspace.textDocuments
        .filter(d => d.fileName.startsWith(this.cache.cacheBasePath))
        .filter(d => !Object.keys(this.zip.files).filter(f => f.includes(basename(d.fileName))))
        .forEach(async t => {
          await window.showTextDocument(Uri.file(t.fileName), { preview: true, preserveFocus: false });
          await commands.executeCommand('workbench.action.closeActiveEditor');
        });
    } catch (err) {
      console.error(err);
    }
  }

  /**
   * @description Check if an OOXML part is different from its cached version
   * @method _hasFileBeenChangedFromOutside
   * @async
   * @private
   * @param  {string} filePath The path of the file
   * @returns {Promise<boolean>}
   */
  private async _hasFileBeenChangedFromOutside(filePath: string, newContent: Uint8Array): Promise<boolean> {
    try {
      const fileArrayBuffer = newContent;
      const prevFileArrayBuffer = await this.cache.getPrevFile(filePath);
      if (Buffer.from(fileArrayBuffer).equals(Buffer.from(prevFileArrayBuffer))) {
        return false;
      }

      const decoder = new TextDecoder();
      const fileText = decoder.decode(fileArrayBuffer);
      const prevFileText = decoder.decode(prevFileArrayBuffer);
      const minFileText = vkBeautify.xmlmin(fileText);
      const minPrevFileText = vkBeautify.xmlmin(prevFileText);
      return !(minFileText === minPrevFileText);
    } catch (err) {
      console.error(err.message || err);
    }

    return false;
  }
}
