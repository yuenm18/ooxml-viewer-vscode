import { exec } from 'child_process';
import { existsSync, mkdir, PathLike, readFile, stat, Stats, writeFile } from 'fs';
import JSZip, { JSZipObject } from 'jszip';
import { basename, dirname, format, join, parse } from 'path';
import rimraf from 'rimraf';
import { promisify } from 'util';
import {
  commands,
  Disposable,
  ExtensionContext,
  FileSystemWatcher,
  Position,
  Range,
  TextDocument,
  TextEditor,
  TextEditorEdit,
  ThemeIcon,
  Uri,
  window,
  workspace,
} from 'vscode';
import formatXml from 'xml-formatter';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';
const execPromise = promisify(exec);
const readFilePromise = promisify(readFile);
const writeFilePromise = promisify(writeFile);
const rimrafPromise = promisify(rimraf);
const statPromise = promisify(stat);
const mkdirPromise = promisify(mkdir);

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  zip: JSZip;
  private _context: ExtensionContext;
  static watchers: Disposable[] = [];
  static watchActions: { [key: string]: number } = {};
  static openTextEditors: { [key: string]: FileNode } = {};
  static cacheFolderName = '.53d3a0ba-37e3-41cf-a068-b10b392cf8ca';
  static ooxmlFilePath: string;
  static fileCachePath: string = join(process.cwd(), OOXMLViewer.cacheFolderName);
  static existsSync = existsSync;
  static execPromise = execPromise;
  static writeFilePromise = writeFilePromise;

  constructor(context: ExtensionContext) {
    this.treeDataProvider = new OOXMLTreeDataProvider(context);
    this.zip = new JSZip();
    this._context = context;
  }

  /**
   * Loads the selected OOXML file into the tree view
   *
   * @param file The OOXML file
   */
  async viewContents(file: Uri): Promise<void> {
    try {
      OOXMLViewer.ooxmlFilePath = file.fsPath;
      await this._resetOOXMLViewer();
      const data = await readFilePromise(file.fsPath);
      await this.zip.loadAsync(data);

      await this._populateOOXMLViewer(this.zip.files, false);
      await mkdirPromise(OOXMLViewer.fileCachePath, { recursive: true });

      const watcher: FileSystemWatcher = workspace.createFileSystemWatcher(file.fsPath);

      watcher.onDidChange((uri: Uri) => {
        this._reloadOoxmlFile(file.fsPath);
      });

      const textDocumentWatcher = workspace.onDidSaveTextDocument(async (e: TextDocument) => {
        try {
          const { fileName } = e;
          let prevFilePath = '';
          if (fileName) {
            prevFilePath = OOXMLViewer._getPrevFilePath(fileName);
          }
          if (fileName && existsSync(fileName) && prevFilePath && existsSync(prevFilePath)) {
            const stats: Stats = await statPromise(fileName);
            const time = stats.mtime.getTime();
            if (!stats.isDirectory() && OOXMLViewer.watchActions[fileName] !== time) {
              OOXMLViewer.watchActions[fileName] = time;
              const data: Buffer = await readFilePromise(fileName);
              const prevData: Buffer = await readFilePromise(prevFilePath);
              if (!data.equals(prevData)) {
                const pathArr = fileName.split(OOXMLViewer.cacheFolderName);
                let normalizedPath: string = pathArr[pathArr.length - 1].replace(/\\/g, '/');
                normalizedPath = normalizedPath.startsWith('/') ? normalizedPath.substring(1) : normalizedPath;
                const zipFile = await this.zip.file(normalizedPath, data, { binary: true }).generateAsync({ type: 'nodebuffer' });
                await writeFilePromise(OOXMLViewer.ooxmlFilePath, zipFile);
                await writeFilePromise(join(dirname(prevFilePath), `compare.${basename(fileName)}`), prevData);
                await writeFilePromise(prevFilePath, data);
              }
            }
          }
        } catch (err) {
          if (err?.code === 'EBUSY') {
            window.showWarningMessage(
              `File not saved.\n${OOXMLViewer.ooxmlFilePath} is open in another program.\nClose that program before making any changes.`,
              { modal: true },
            );
            OOXMLViewer._makeDirty(window.activeTextEditor);
          }
        }
      });

      const closeWatcher = workspace.onDidCloseTextDocument((textDocument: TextDocument) => {
        delete OOXMLViewer.openTextEditors[textDocument.fileName];
      });
      OOXMLViewer.watchers.push(watcher, textDocumentWatcher, closeWatcher);
    } catch (err) {
      console.error(err);
      window.showErrorMessage(`Could not load ${file.fsPath}`, err);
    }
  }

  /**
   * Displays the selected file
   *
   * @param fileNode The selected file node
   */
  async viewFile(fileNode: FileNode, preview = true): Promise<void> {
    try {
      const folderPath = join(OOXMLViewer.fileCachePath, dirname(fileNode.fullPath));
      const filePath: string = join(folderPath, fileNode.fileName);
      await this._createFile(fileNode.fullPath, fileNode.fileName);
      const uri: Uri = Uri.parse(`file:///${filePath}`);
      OOXMLViewer.openTextEditors[filePath] = fileNode;
      const xmlDoc: TextDocument = await workspace.openTextDocument(uri);

      await window.showTextDocument(xmlDoc, {preview});
    } catch (e) {
      console.error(e);
      window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
    }
  }

  /**
   * Clears the OOXML viewer
   */
  clear(): Promise<void> {
    return this._resetOOXMLViewer();
  }

  /**
   * Compares the file to it's previous version
   */
  async getDiff(file: FileNode): Promise<void> {
    try {
      const filePath = join(OOXMLViewer.fileCachePath, dirname(file.fullPath), file.fileName);
      const compareFilePath = join(OOXMLViewer.fileCachePath, dirname(file.fullPath), `compare.${file.fileName}`);
      const rightUri = Uri.file(filePath);
      const leftUri = Uri.file(compareFilePath);
      const title = `${basename(filePath)} ↔ ${basename(compareFilePath)}`;
      await commands.executeCommand('vscode.diff', leftUri, rightUri, title);
    } catch (err) {
      console.error(err);
    }
  }

  private async _viewFiles(fileNodes: FileNode[]): Promise<void> {
    while (fileNodes.length) {
      const fileNode: FileNode | undefined = fileNodes.pop();
      if (fileNode) {
        await this.viewFile(fileNode, false);
      }
    }
  }
  private static async _closeEditors(textDocuments?: TextDocument[]): Promise<void> {
    const tds =
      textDocuments ?? workspace.textDocuments.filter(t => t.fileName.toLowerCase().includes(OOXMLViewer.fileCachePath.toLowerCase()));
    if (tds.length) {
      const td: TextDocument | undefined = tds.pop();
      if (td) {
        await window.showTextDocument(td, { preview: true, preserveFocus: false });
        await commands.executeCommand('workbench.action.closeActiveEditor');
        await OOXMLViewer._closeEditors(tds);
      }
    }
  }

  private async _resetOOXMLViewer(): Promise<void> {
    try {
      this.zip = new JSZip();
      this.treeDataProvider.rootFileNode = new FileNode(this._context);
      this.treeDataProvider.refresh();
      if (OOXMLViewer.existsSync(OOXMLViewer.fileCachePath)) {
        await rimrafPromise(OOXMLViewer.fileCachePath);
      }
      await OOXMLViewer.closeWatchers();
      await OOXMLViewer._closeEditors();
    } catch (err) {
      console.error(err);
      window.showErrorMessage('Could not remove ooxml file viewer cache');
    }
  }

  private async _populateOOXMLViewer(files: { [key: string]: JSZip.JSZipObject }, showNewFileLabel: boolean) {
    for (const fileWithPath of Object.keys(files)) {
      // ignore folder files
      if (files[fileWithPath].dir) {
        continue;
      }

      // Build nodes for each file
      let currentFileNode = this.treeDataProvider.rootFileNode;
      for (const fileOrFolderName of fileWithPath.split('/')) {
        // Create node if it does not exist
        const existingFileNode = currentFileNode.children.find(c => c.description === fileOrFolderName);
        if (existingFileNode) {
          const warningIcon: string = this._context.asAbsolutePath(join('images', 'asterisk.svg'));
          currentFileNode = existingFileNode;
          await this._createFile(currentFileNode.fullPath, currentFileNode.fileName);
          const filesAreDifferent = await OOXMLViewer._fileHasBeenChangedFromOutside(currentFileNode.fullPath);
          if (filesAreDifferent) {
            currentFileNode.iconPath = warningIcon;
            const path = OOXMLViewer._getPrevFilePath(currentFileNode.fullPath);
            await this._createFile(path, `compare.${currentFileNode.fileName}`);
            await this._createFile(currentFileNode.fullPath, `prev.${currentFileNode.fileName}`);
          } else {
            currentFileNode.iconPath = currentFileNode.children.length ? ThemeIcon.Folder : ThemeIcon.File;
          }
        } else {
          const newFileNode = new FileNode(this._context);
          newFileNode.fileName = fileOrFolderName;
          newFileNode.parent = currentFileNode;
          newFileNode.fullPath = fileWithPath;
          currentFileNode.children.push(newFileNode);
          currentFileNode = newFileNode;
          await this._createFile(newFileNode.fullPath, newFileNode.fileName);
          await this._createFile(newFileNode.fullPath, `prev.${newFileNode.fileName}`);
          if (showNewFileLabel) {
            const warningIconGreen: string = this._context.asAbsolutePath(join('images', 'asterisk.green.svg'));
            const compareFilePath: PathLike = join(
              OOXMLViewer.fileCachePath,
              dirname(newFileNode.fullPath),
              `compare.${newFileNode.fileName}`,
            );
            currentFileNode.iconPath = warningIconGreen;
            await writeFilePromise(compareFilePath, '');
          } else {
            await this._createFile(newFileNode.fullPath, `compare.${newFileNode.fileName}`);
          }
        }
      }

      currentFileNode.fullPath = fileWithPath;
    }

    this.treeDataProvider.refresh();
  }

  private async _createFile(fullPath: string, fileName: string): Promise<void> {
    try {
      const folderPath = join(OOXMLViewer.fileCachePath, dirname(fullPath));
      const filePath: string = join(folderPath, fileName);
      await mkdirPromise(folderPath, { recursive: true });
      if (process.platform.startsWith('win')) {
        const { stderr } = await OOXMLViewer.execPromise('attrib +h ' + OOXMLViewer.fileCachePath);
        if (stderr) {
          throw new Error(stderr);
        }
      }
      const file: JSZipObject | null = this.zip.file(fullPath);
      const text: string = await file?.async('text') ?? await (await readFilePromise(preFilePath)).toString();
      if (text.startsWith('<?xml')) {
        let formattedXml = '';
        if (text.length < 100000){
          formattedXml = formatXml(text);
        }
        await OOXMLViewer.writeFilePromise(filePath, formattedXml || text, 'utf8');
      }
    } catch (err) {
      console.error(err);
    }
  }

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

  private static _getPrevFilePath(path: string): string {
    const { dir, base } = parse(path);
    return format({ dir, base: `prev.${base}` });
  }

  static async closeWatchers(): Promise<void> {
    if (OOXMLViewer.watchers.length) {
      OOXMLViewer.watchers.forEach(w => w.dispose());
      OOXMLViewer.watchers = [];
    }
  }

  private async _reloadOoxmlFile(filePath: string): Promise<void> {
    const ooxmlZip: JSZip = new JSZip();
    const data: Buffer = await readFilePromise(filePath);
    await ooxmlZip.loadAsync(data);
    this.zip = ooxmlZip;
    this._populateOOXMLViewer(this.zip.files, true);
    this._viewFiles(Object.values(OOXMLViewer.openTextEditors));
  }

  private static async _fileHasBeenChangedFromOutside(firstFile: string): Promise<boolean> {
    try {
      const secondFile = OOXMLViewer._getPrevFilePath(firstFile);
      const firstFilePath = join(OOXMLViewer.fileCachePath, firstFile);
      const secondFilePath = join(OOXMLViewer.fileCachePath, secondFile);
      const firstStat: Stats = await statPromise(firstFilePath);
      const secondStat: Stats = await statPromise(secondFilePath);
      if (!firstStat.isDirectory() && !secondStat.isDirectory()) {
        const firstBuffer: Buffer = await readFilePromise(firstFilePath);
        const secondBuffer: Buffer = await readFilePromise(secondFilePath);
        if (!firstBuffer.equals(secondBuffer)) {
          return true;
        }
      }
    } catch (err) {
      console.error(err.message || err);
    }
    return false;
  }
}
