import { existsSync, PathLike } from 'fs';
import JSZip, { JSZipObject } from 'jszip';
import { basename, dirname, join, normalize, parse } from 'path';
import vkBeautify from 'vkbeautify';
import {
  commands,
  Disposable,
  ExtensionContext,
  FileStat,
  FileSystemWatcher,
  FileType,
  Position,
  ProgressLocation,
  Range,
  TextDocument,
  TextEditor,
  TextEditorEdit,
  ThemeIcon,
  Uri,
  window,
  workspace,
} from 'vscode';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';
/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  zip: JSZip;
  static watchers: Disposable[] = [];
  static watchActions: { [key: string]: number } = {};
  static openTextEditors: { [key: string]: FileNode } = {};
  static cacheFolderName = '.open-xml-viewer';
  static ooxmlFilePath: string;
  fileCachePath: string = join(this._context.storageUri?.fsPath || '', OOXMLViewer.cacheFolderName);
  /**
   * Constructs an instance of OOXMLViewer
   * @constructor OOXMLViewer
   * @param  ExtensionContext _context
   * @returns {OOXMLViewer} instance
   */
  constructor(private _context: ExtensionContext) {
    this.treeDataProvider = new OOXMLTreeDataProvider();
    this.zip = new JSZip();
  }

  /**
   * Loads the selected OOXML file into the tree view
   * and add file listeners
   * @method viewContents
   * @async
   * @param file The OOXML file
   * @returns {Promise} Promise that returns void
   */
  async viewContents(file: Uri): Promise<void> {
    try {
      OOXMLViewer.ooxmlFilePath = file.fsPath;
      await window.withProgress(
        {
          location: ProgressLocation.Notification,
          title: 'OOXML Viewer',
        },
        async progress => {
          progress.report({ message: 'Unpacking OOXML Parts' });
          await this._resetOOXMLViewer();
          const data = await workspace.fs.readFile(Uri.file(file.fsPath));
          await this.zip.loadAsync(data);
          await this._populateOOXMLViewer(this.zip.files, false);
          await workspace.fs.createDirectory(Uri.file(this.fileCachePath));

          const watcher: FileSystemWatcher = workspace.createFileSystemWatcher(file.fsPath);

          watcher.onDidChange((uri: Uri) => {
            this._reloadOoxmlFile(file.fsPath);
          });

          const textDocumentWatcher = workspace.onDidSaveTextDocument(this._updateOOXMLFile);
          // TODO: find a better way to remove closed text editors from the openTextEditors. The onDidCloseTextDocument takes more than 3+ minutes to fire.
          const closeWatcher = workspace.onDidCloseTextDocument((textDocument: TextDocument) => {
            delete OOXMLViewer.openTextEditors[textDocument.fileName];
          });
          OOXMLViewer.watchers.push(watcher, textDocumentWatcher, closeWatcher);
        },
      );
    } catch (err) {
      console.error(err);
      window.showErrorMessage(`Could not load ${file.fsPath}`, err);
    }
  }

  /**
   * Displays the selected file
   * @async
   * @method viewFile
   * @param fileNode The selected file node
   * @returns {Promise} Promise that returns void
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
          const folderPath = join(this.fileCachePath, dirname(fileNode.fullPath));
          const filePath: string = join(folderPath, fileNode.fileName);
          await this._createFile(fileNode.fullPath, fileNode.fileName, true);
          const uri: Uri = Uri.parse(`file:///${filePath}`);
          OOXMLViewer.openTextEditors[filePath] = fileNode;
          commands.executeCommand('vscode.open', uri);
        },
      );
    } catch (e) {
      console.error(e);
      window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
    }
  }

  /**
   * Clears the OOXML viewer
   * @method clear
   * @returns {Promise} Promise that returns void
   */
  clear(): Promise<void> {
    return this._resetOOXMLViewer();
  }

  /**
   * @method getDiff
   * @async
   * @param  {FileNode} file the FileNode of the file to be diffed
   * @returns {Promise} Promise that returns void
   * @description Opens tab showing the difference between the just the primary xml part and the compare xml part
   */
  async getDiff(file: FileNode): Promise<void> {
    try {
      // get the full path for the primary file and the compare files
      const filePath = join(this.fileCachePath, dirname(file.fullPath), file.fileName);
      const compareFilePath = join(this.fileCachePath, dirname(file.fullPath), `compare.${file.fileName}`);
      // create URIs and title
      const enc = new TextEncoder();
      const rightUri = Uri.file(filePath);
      const leftUri = Uri.file(compareFilePath);
      const rightXml = await (await workspace.fs.readFile(rightUri)).toString();
      const leftXml = await (await workspace.fs.readFile(leftUri)).toString();
      await workspace.fs.writeFile(rightUri, enc.encode(rightXml.startsWith('<?xml') ? vkBeautify.xml(rightXml) : rightXml));
      await workspace.fs.writeFile(leftUri, enc.encode(leftXml.startsWith('<?xml') ? vkBeautify.xml(leftXml) : leftXml));
      const title = `${basename(filePath)} ↔ ${basename(compareFilePath)}`;
      // diff the primary and compare files
      await commands.executeCommand('vscode.diff', leftUri, rightUri, title);
    } catch (err) {
      console.error(err);
    }
  }
  /**
   * Close all file watchers
   * @method closeWatchers
   * @async
   * @returns Promise
   */
  static closeWatchers(): void {
    if (OOXMLViewer.watchers.length) {
      OOXMLViewer.watchers.forEach(w => w.dispose());
      OOXMLViewer.watchers = [];
    }
  }
  /**
   * Receives an array of FileNode instances and calls viewFile on each and empties the array
   * @method _viewFiles
   * @private
   * @async
   * @param  {FileNode[]} fileNodes
   * @returns Promise
   */
  private async _viewFiles(fileNodes: FileNode[]): Promise<void> {
    while (fileNodes.length) {
      const fileNode: FileNode | undefined = fileNodes.pop();
      const filePath: string | undefined = fileNode ? join(this.fileCachePath, dirname(fileNode.fullPath), fileNode.fileName) : undefined;
      if (fileNode && filePath && existsSync(filePath)) {
        await this.viewFile(fileNode);
      }
    }
  }
  /**
   * Closes all active editor tabs
   * @method _closeEditors
   * @private
   * @async
   * @param  {TextDocument[]} textDocuments?
   * @returns Promise
   */
  async closeEditors(textDocuments?: TextDocument[]): Promise<void> {
    try {
      const tds = textDocuments ?? workspace.textDocuments.filter(t => t.fileName.toLowerCase().includes(this.fileCachePath.toLowerCase()));
      if (tds.length) {
        const td: TextDocument | undefined = tds.pop();
        if (td) {
          await window.showTextDocument(td, { preview: true, preserveFocus: false });
          await commands.executeCommand('workbench.action.closeActiveEditor');
          await this.closeEditors(tds);
        }
      }
    } catch (err) {
      console.error(err);
    }
  }
  /**
   * Sets this.zip to an empty zip file, deletes the cache folder, closes all watchers, and closes all editor tabs
   * @method _resetOOXMLViewer
   * @private
   * @async
   * @returns Promise
   */
  private async _resetOOXMLViewer(): Promise<void> {
    try {
      this.zip = new JSZip();
      this.treeDataProvider.rootFileNode = new FileNode();
      this.treeDataProvider.refresh();
      if (existsSync(this.fileCachePath)) {
        // do not await this
        workspace.fs.delete(Uri.file(this.fileCachePath), { recursive: true, useTrash: false });
      }
      OOXMLViewer.closeWatchers();
      await this.closeEditors();
    } catch (err) {
      console.error(err);
      window.showErrorMessage('Could not remove ooxml file viewer cache');
    }
  }
  /**
   * Create or update tree view File Nodes and create cache files for comparison
   * @method _populateOOXMLViewer
   * @private
   * @async
   * @param  {{[key:string]:JSZip.JSZipObject}} files
   * @param  {boolean} showNewFileLabel
   */
  private async _populateOOXMLViewer(files: { [key: string]: JSZip.JSZipObject }, showNewFileLabel: boolean) {
    const fileKeys: string[] = Object.keys(files);
    for (const fileWithPath of fileKeys) {
      // ignore folder files
      if (files[fileWithPath].dir) {
        continue;
      }

      // Build nodes for each file
      let currentFileNode = this.treeDataProvider.rootFileNode;
      const names: string[] = fileWithPath.split('/');
      for (const fileOrFolderName of names) {
        // Create node if it does not exist
        const existingFileNode = currentFileNode.children.find(c => c.description === fileOrFolderName);
        if (existingFileNode) {
          const warningIcon: string = this._context.asAbsolutePath(join('images', 'asterisk.yellow.svg'));
          currentFileNode = existingFileNode;
          await this._createFile(currentFileNode.fullPath, currentFileNode.fileName);
          const filesAreDifferent = await this._fileHasBeenChangedFromOutside(currentFileNode.fullPath);
          // If the files are different replace copy prev to compare and copy current part to prev and add the warning icon,
          // else don't edit the files and use a folder or file icon
          const path = OOXMLViewer._getPrevFilePath(currentFileNode.fullPath);
          const comparePath = `compare.${basename(currentFileNode.fileName).replace('compare.', '')}`;
          const prevPath = `prev.${basename(currentFileNode.fileName).replace('prev.', '')}`;
          await this._createFile(path, comparePath);
          await this._createFile(currentFileNode.fullPath, prevPath);
          if (filesAreDifferent) {
            currentFileNode.iconPath = Uri.file(warningIcon);
          } else {
            currentFileNode.iconPath = currentFileNode.children.length ? ThemeIcon.Folder : ThemeIcon.File;
          }
        } else {
          // create a new FileNode with the currentFileNode as parent and add it to the currentFileNode children
          const newFileNode = new FileNode();
          newFileNode.fileName = fileOrFolderName;
          newFileNode.parent = currentFileNode;
          newFileNode.fullPath = fileWithPath;
          currentFileNode.children.push(newFileNode);
          currentFileNode = newFileNode;
          // create a copy of the file and a prev copy with the same data i.e. no changes have been made
          await this._createFile(newFileNode.fullPath, newFileNode.fileName);
          await this._createFile(newFileNode.fullPath, `prev.${basename(newFileNode.fileName).replace('prev.', '')}`);
          // if showNewFileLabel is true (meaning this is a refresh of the tree) then this is a new file so
          // add a green asterisk as the icon and create an empty compare file
          // else create the compare file from the newFileNode.fullPath
          if (showNewFileLabel) {
            const warningIconGreen: string = this._context.asAbsolutePath(join('images', 'asterisk.green.svg'));
            const compareFilePath: PathLike = join(this.fileCachePath, dirname(newFileNode.fullPath), `compare.${newFileNode.fileName}`);
            currentFileNode.iconPath = Uri.file(warningIconGreen);
            await workspace.fs.writeFile(Uri.file(compareFilePath), new Uint8Array());
          } else {
            await this._createFile(newFileNode.fullPath, `compare.${basename(newFileNode.fileName).replace(/compare.|prev./, '')}`);
          }
        }
      }
      // set the current node fullPath (either new or existing) to fileWithPath
      currentFileNode.fullPath = fileWithPath;
    }
    await this._removeDeletedParts();
    await this._reformatOpenTabs();
    // tell vscode the tree has changed
    this.treeDataProvider.refresh();
  }
  /**
   * Create a file from a part of the zip file
   * @method _createFile
   * @private
   * @async
   * @param  {string} relativePath the path to the folder in the zip file
   * @param  {string} fileName the name of the file
   * @returns Promise
   */
  private async _createFile(relativePath: string, fileName: string, formatIt = false): Promise<void> {
    try {
      const preFilePath = join(this.fileCachePath, relativePath);
      const folderPath = join(this.fileCachePath, dirname(relativePath));
      const filePath: string = join(folderPath, fileName);
      await workspace.fs.createDirectory(Uri.file(folderPath));
      const file: JSZipObject | null = this.zip.file(relativePath);
      const text: string = (await file?.async('text')) ?? (await (await workspace.fs.readFile(Uri.file(preFilePath))).toString());
      if (text.startsWith('<?xml')) {
        let formattedXml = '';
        const len = text.replace(/\s+/g, '').length;
        if (formatIt && len < 1000000) {
          formattedXml = vkBeautify.xml(text);
        }
        if (len >= 1000000 && formatIt) {
          window.showWarningMessage(
            `${basename(relativePath)} is too large to format.\nOOXML Parts must be less than 1,000,000 characters to format`,
            { modal: true },
          );
        }
        const enc = new TextEncoder();
        await workspace.fs.writeFile(Uri.file(filePath), enc.encode(formattedXml || text));
        if (!/compare|prev/.test(filePath)) {
          await this.zip.file(relativePath, text);
        }
      } else {
        const u8a: Uint8Array | undefined = await file?.async('uint8array');
        if (u8a) {
          await workspace.fs.writeFile(Uri.file(filePath), u8a);
        }
      }
    } catch (err) {
      console.error(err);
    }
  }
  /**
   * Writes changes to OOXML file being inspected
   * @method _updateOOXMLFile
   * @async
   * @private
   * @param  {TextDocument} e
   * @returns Promise
   */
  private _updateOOXMLFile = async (e: TextDocument): Promise<void> => {
    try {
      const { fileName } = e;
      let prevFilePath = '';
      if (fileName) {
        prevFilePath = OOXMLViewer._getPrevFilePath(fileName);
      }
      if (fileName && existsSync(fileName) && prevFilePath && existsSync(prevFilePath)) {
        const stats: FileStat = await workspace.fs.stat(Uri.file(fileName));
        const time = stats.mtime;
        if (stats.type !== FileType.Directory && OOXMLViewer.watchActions[fileName] !== time) {
          OOXMLViewer.watchActions[fileName] = time;
          const textDecoder = new TextDecoder();
          const cur = await workspace.fs.readFile(Uri.file(fileName));
          const prev = await workspace.fs.readFile(Uri.file(prevFilePath));
          const curMiniXml = vkBeautify.xmlmin(textDecoder.decode(cur), true);
          const prevMiniXml = vkBeautify.xmlmin(textDecoder.decode(prev), true);
          const data: Buffer = Buffer.from(curMiniXml);
          const prevData: Buffer = Buffer.from(prevMiniXml);
          const checkName = basename(fileName);
          if (!data.equals(prevData)) {
            const pathArr = fileName.split(OOXMLViewer.cacheFolderName);
            let normalizedPath: string = pathArr[pathArr.length - 1].replace(/\\/g, '/');
            normalizedPath = normalizedPath.startsWith('/') ? normalizedPath.substring(1) : normalizedPath;
            const zipFile = await this.zip.file(normalizedPath, data, { binary: true }).generateAsync({ type: 'uint8array' });
            await workspace.fs.writeFile(Uri.file(OOXMLViewer.ooxmlFilePath), zipFile);
            await workspace.fs.writeFile(Uri.file(join(dirname(prevFilePath), `compare.${checkName}`)), prev);
            await workspace.fs.writeFile(Uri.file(prevFilePath), cur);
            this.treeDataProvider.refresh();
          }
        }
      }
    } catch (err) {
      if (err?.code === 'EBUSY' || err?.message.toLowerCase().includes('ebusy')) {
        window.showWarningMessage(
          `File not saved.\n${basename(OOXMLViewer.ooxmlFilePath)} is open in another program.\n
                  Close that program before making any changes.`,
          { modal: true },
        );
        OOXMLViewer._makeDirty(window.activeTextEditor);
      }
    }
  };
  /**
   * Make a text editor tab dirty
   * @method _makeDirty
   * @async
   * @private
   * @param  {TextEditor} activeTextEditor? the text editor to be made dirty
   * @returns Promise
   */
  private static async _makeDirty(activeTextEditor?: TextEditor): Promise<void> {
    activeTextEditor?.edit(async (textEditorEdit: TextEditorEdit) => {
      if (window.activeTextEditor?.selection) {
        const { activeTextEditor } = window;

        if (activeTextEditor && activeTextEditor.document.lineCount >= 2) {
          const lineNumber = activeTextEditor.document.lineCount - 2;
          const lastLineRange = new Range(new Position(lineNumber, 0), new Position(lineNumber + 1, 0));
          const lastLineText = activeTextEditor.document.getText(lastLineRange);
          textEditorEdit.replace(lastLineRange, lastLineText);
          return;
        }

        // Try to replace the first character.
        const range = new Range(new Position(0, 0), new Position(0, 1));
        const text: string | undefined = activeTextEditor?.document.getText(range);
        if (text) {
          textEditorEdit.replace(range, text);
          return;
        }

        // With an empty file, we first add a character and then remove it.
        // This has to be done as two edits, which can cause the cursor to
        // visibly move and then return, but we can at least combine them
        // into a single undo step.
        await activeTextEditor?.edit(
          (innerEditBuilder: TextEditorEdit) => {
            innerEditBuilder.replace(range, ' ');
          },
          { undoStopBefore: true, undoStopAfter: false },
        );

        await activeTextEditor?.edit(
          (innerEditBuilder: TextEditorEdit) => {
            innerEditBuilder.replace(range, '');
          },
          { undoStopBefore: false, undoStopAfter: true },
        );
      }
    });
  }
  /**
   * Get the file path for the prev version of a file
   * @method _getPrevFilePath
   * @private
   * @async
   * @param  {string} path
   * @returns string
   */
  private static _getPrevFilePath(path: string): string {
    const { dir, base } = parse(path);
    return normalize(`${dir}/prev.${base}`);
  }
  /**
   * Update the OOXML cache files
   * @method _reloadOoxmlFile
   * @async
   * @private
   * @param  {string} filePath Path to the OOXML file to load
   * @returns Promise
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
          const ooxmlZip: JSZip = new JSZip();
          const data: Buffer = Buffer.from(await workspace.fs.readFile(Uri.file(filePath)));
          await ooxmlZip.loadAsync(data);
          this.zip = ooxmlZip;
          await this._populateOOXMLViewer(this.zip.files, true);
        } catch (err) {
          console.error(err);
        }
      },
    );
  }
  /**
   * Delete cache files for parts deleted from OOXML file
   * @method _removeDeletedParts
   * @async
   * @private
   * @param  {FileNode} node?
   * @returns Promise
   */
  private async _removeDeletedParts(node?: FileNode): Promise<void> {
    try {
      const fileNames = Object.keys(this.zip.files);
      if (!node) {
        await this._removeDeletedParts(this.treeDataProvider.rootFileNode);
      } else {
        node.children.forEach(async (n, i, arr) => {
          const path = join(this.fileCachePath, n.fullPath);
          if (!fileNames.includes(n.fullPath)) {
            const file: string = await (await workspace.fs.readFile(Uri.file(path))).toString();
            if (file) {
              n.iconPath = Uri.file(this._context.asAbsolutePath(join('images', 'asterisk.red.svg')));
              await workspace.fs.writeFile(Uri.file(join(this.fileCachePath, n.fullPath)), new Uint8Array());
              this.treeDataProvider.refresh();
            } else {
              workspace.fs.delete(Uri.file(path), { recursive: true, useTrash: false });
              workspace.fs.delete(Uri.file(join(dirname(path), `prev.${basename(path)}`)), { recursive: true, useTrash: false });
              workspace.fs.delete(Uri.file(join(dirname(path), `compare.${basename(path)}`)), { recursive: true, useTrash: false });
              arr.splice(i, 1);
              this.treeDataProvider.refresh();
            }
          } else if (n.children.length) {
            n.children.forEach(async nn => await this._removeDeletedParts(nn));
          }
        });
      }
    } catch (err) {
      console.error(err);
    }
  }
  private async _reformatOpenTabs(): Promise<void> {
    try {
      const textDocuments: TextDocument[] = [];
      const closedTextDocuments: TextDocument[] = [];
      workspace.textDocuments.forEach(t => {
        if (t.fileName.toLowerCase().includes(this.fileCachePath.toLowerCase())) {
          if (Object.keys(this.zip.files).filter(f => f.includes(basename(t.fileName).replace(/compare.|prev./, ''))).length) {
            textDocuments.push(t);
          } else {
            closedTextDocuments.push(t);
          }
        }
      });
      textDocuments.forEach(async td => {
        try {
          const prevFileName = join(dirname(td.fileName), `prev.${basename(td.fileName.replace(/compare.|prev./, ''))}`);
          const compareFileName = join(dirname(td.fileName), `compare.${basename(td.fileName.replace(/compare.|prev./, ''))}`);
          const xml = await (await workspace.fs.readFile(Uri.file(td.fileName))).toString();
          const prevXml = await (await workspace.fs.readFile(Uri.file(prevFileName))).toString();
          const compareXml = await (await workspace.fs.readFile(Uri.file(compareFileName))).toString();
          const enc = new TextEncoder();
          [
            { xml, fileName: td.fileName },
            { xml: prevXml, fileName: prevFileName },
            { xml: compareXml, fileName: compareFileName },
          ].forEach(async ({ xml, fileName }) => {
            if (xml.startsWith('<?xml')) {
              const text = vkBeautify.xml(xml);
              await workspace.fs.writeFile(Uri.file(fileName), enc.encode(text));
            }
          });
        } catch (err) {
          console.error(err);
        }
      });
      closedTextDocuments.forEach(async t => {
        await window.showTextDocument(Uri.file(t.fileName), { preview: true, preserveFocus: false });
        await commands.executeCommand('workbench.action.closeActiveEditor');
      });
    } catch (err) {
      console.error(err);
    }
  }
  /**
   * Check if an OOXML part is different from its cached version
   * @method _fileHasBeenChangedFromOutside
   * @async
   * @private
   * @param  {string} firstFile
   * @returns Promise
   */
  private async _fileHasBeenChangedFromOutside(firstFile: string): Promise<boolean> {
    try {
      const secondFile = OOXMLViewer._getPrevFilePath(firstFile);
      const firstFilePath = join(this.fileCachePath, firstFile);
      const secondFilePath = join(this.fileCachePath, secondFile);
      const firstStat: FileStat = await workspace.fs.stat(Uri.file(firstFilePath));
      const secondStat: FileStat = await workspace.fs.stat(Uri.file(secondFilePath));
      if (firstStat.type !== FileType.Directory && secondStat.type !== FileType.Directory) {
        const firstBuffer: Buffer = Buffer.from(await workspace.fs.readFile(Uri.file(firstFilePath)));
        const secondBuffer: Buffer = Buffer.from(await workspace.fs.readFile(Uri.file(secondFilePath)));
        if (!firstBuffer.equals(secondBuffer)) {
          const decoder = new TextDecoder();
          const first = decoder.decode(await workspace.fs.readFile(Uri.file(firstFilePath)));
          const second = decoder.decode(await workspace.fs.readFile(Uri.file(secondFilePath)));
          const firstBufferMin: Buffer = Buffer.from(vkBeautify.xmlmin(first));
          const secondBufferMin: Buffer = Buffer.from(vkBeautify.xmlmin(second));
          return !firstBufferMin.equals(secondBufferMin);
        }
      }
    } catch (err) {
      console.error(err.message || err);
    }
    return false;
  }
}
