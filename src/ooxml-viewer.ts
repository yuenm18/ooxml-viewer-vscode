import { exec } from 'child_process';
import { existsSync, readFile, stat, Stats, writeFile } from 'fs';
import JSZip, { JSZipObject } from 'jszip';
import mkdirp from 'mkdirp';
import { dirname, format, join, parse } from 'path';
import rimraf from 'rimraf';
import { promisify } from 'util';
import vscode, { FileSystemWatcher, Position, TextDocument, TextEditor, TextEditorEdit, Uri } from 'vscode';
import formatXml from 'xml-formatter';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';
const execPromise = promisify(exec);
const readFilePromise = promisify(readFile);
const writeFilePromise = promisify(writeFile);
const rimrafPromise = promisify(rimraf);
const statPromise = promisify(stat);

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  zip: JSZip;
  static watchers: FileSystemWatcher[] = [];
  static watchActions: { [key: string]: number; } = {};
  static openTextEditors: { [key: string]: FileNode; } = {};
  static cacheFolderName = '.ooxml-temp-file-folder-78kIPsmTq5TK';
  static ooxmlFilePath: string;
  static rootPath: string = vscode.workspace.rootPath == undefined ? parse(process.cwd()).root : vscode.workspace.rootPath;
  static fileCachePath: string = join(OOXMLViewer.rootPath, OOXMLViewer.cacheFolderName);
  static existsSync = existsSync;
  static mkdirp = mkdirp;
  static execPromise = execPromise;
  static writeFilePromise = writeFilePromise;

  constructor() {
    this.treeDataProvider = new OOXMLTreeDataProvider();
    this.zip = new JSZip();
  }

  /**
   * Loads the selected OOXML file into the tree view
   *
   * @param file The OOXML file
   */
  async viewContents(file: vscode.Uri): Promise<void> {
    try {
      OOXMLViewer.ooxmlFilePath = file.fsPath;
      await this.resetOOXMLViewer();
      const data = await readFilePromise(file.fsPath);
      await this.zip.loadAsync(data);
      await this.populateOOXMLViewer(this.zip.files);
      await OOXMLViewer.mkdirp(OOXMLViewer.fileCachePath);

      const watcher: FileSystemWatcher = vscode.workspace.createFileSystemWatcher(file.fsPath);
      watcher.onDidChange((uri: Uri) => {
        this.checkDiff(file.fsPath);
      });
      OOXMLViewer.watchers.push(watcher);
      vscode.workspace.onDidSaveTextDocument(async (e: TextDocument) => {
        try {
          const { fileName } = e;
          let prevFilePath = '';
          if (fileName) {
            const { dir, base } = parse(fileName);
            prevFilePath = format({ dir, base: `prev.${base}` });
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
                await writeFilePromise(prevFilePath, data);
              }
            }
          }
        } catch (err) {
          if (err?.code === 'EBUSY') {
            vscode.window.showWarningMessage(
              `File not saved.\n${OOXMLViewer.ooxmlFilePath} is open in another program.\nClose that program before making any changes.`,
              { modal: true },
            );
            OOXMLViewer.makeDirty(vscode.window.activeTextEditor);
          }
        }
      });
    } catch (err) {
      console.error(err);
      vscode.window.showErrorMessage(`Could not load ${file.fsPath}`, err);
    }
  }

  /**
   * Displays the selected file
   *
   * @param fileNode The selected file node
   */
  async viewFile(fileNode: FileNode): Promise<void> {
    try {
      const folderPath = join(OOXMLViewer.fileCachePath, dirname(fileNode.fullPath));
      const filePath: string = join(folderPath, fileNode.fileName);
      await this.createFile(fileNode.fullPath, fileNode.fileName);
      const uri: Uri = Uri.parse('file:///' + filePath);
      OOXMLViewer.openTextEditors[filePath] = fileNode;
      const xmlDoc: TextDocument = await vscode.workspace.openTextDocument(uri);

      await vscode.window.showTextDocument(xmlDoc, {preview: false});
    } catch (e) {
      console.error(e);
      vscode.window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
    }
  }

  private async viewFiles(uris: FileNode[]): Promise<void> {
    while (uris.length) {
      const uri: FileNode | undefined = uris.pop();
      if (uri) {
        this.viewFile(uri);
      }

    }

  }

  /**
   * Clears the OOXML viewer
   */
  clear(): Promise<void> {
    return this.resetOOXMLViewer();
  }
  private static async closeEditors(textDocuments?: TextDocument[]): Promise<void> {
    const tds = textDocuments ??
      vscode.workspace.textDocuments.filter(t => t.fileName.toLowerCase().includes(OOXMLViewer.fileCachePath.toLowerCase()));
    if (tds.length) {
      const td: TextDocument | undefined = tds.pop();
      if (td) {
        await vscode.window.showTextDocument(td, { preview: true, preserveFocus: false });
        await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
        await OOXMLViewer.closeEditors(tds);
      }
    }
  }

  private async resetOOXMLViewer(): Promise<void> {
    try {
      this.zip = new JSZip();
      this.treeDataProvider.rootFileNode = new FileNode();
      this.treeDataProvider.refresh();
      if (OOXMLViewer.existsSync(OOXMLViewer.fileCachePath)) {
        await rimrafPromise(OOXMLViewer.fileCachePath);
      }
      await OOXMLViewer.closeWatchers();
      await OOXMLViewer.closeEditors();
    } catch (err) {
      console.error(err);
      vscode.window.showErrorMessage('Could not remove ooxml file viewer cache');
    }
  }

  private async populateOOXMLViewer(files: { [key: string]: JSZip.JSZipObject; }) {
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
          currentFileNode = existingFileNode;
        } else {
          const newFileNode = new FileNode();
          newFileNode.fileName = fileOrFolderName;
          newFileNode.parent = currentFileNode;
          newFileNode.fullPath = fileWithPath;
          currentFileNode.children.push(newFileNode);
          currentFileNode = newFileNode;
        }
      }

      currentFileNode.fullPath = fileWithPath;
    }

    this.treeDataProvider.refresh();
  }

  private async createFile(fullPath: string, fileName: string): Promise<void> {
    try {
      const folderPath = join(OOXMLViewer.fileCachePath, dirname(fullPath));
      const filePath: string = join(folderPath, fileName);
      await OOXMLViewer.mkdirp(folderPath);
      // On Windows hide the folder
      if (process.platform.startsWith('win')) {
        const { stderr } = await OOXMLViewer.execPromise('attrib +h ' + OOXMLViewer.fileCachePath);
        if (stderr) {
          throw new Error(stderr);
        }
      }
      const file: JSZipObject | null = this.zip.file(fullPath);
      const text: string = await file?.async('text') ?? '';
      if (text.startsWith('<?xml')) {
        const formattedXml: string = formatXml(text);
        await OOXMLViewer.writeFilePromise(filePath, formattedXml, 'utf8');
        let prevFilePath = '';
        const { dir, base } = parse(filePath);
        prevFilePath = format({ dir, base: `prev.${base}` });
        if (!existsSync(prevFilePath)) {
          await OOXMLViewer.writeFilePromise(prevFilePath, formattedXml, 'utf8');
        }
      }
    } catch (err) {
      console.error(err);
    }
  }

  private static async makeDirty(activeTextEditor?: TextEditor): Promise<void> {
    activeTextEditor?.edit(async (textEditorEdit: TextEditorEdit) => {
      if (vscode.window.activeTextEditor?.selection) {
        const { activeTextEditor } = vscode.window;

        if (activeTextEditor && activeTextEditor.document.lineCount >= 2) {
          const lineNumber = activeTextEditor.document.lineCount - 2;
          const lastLineRange = new vscode.Range(
            new Position(lineNumber, 0),
            new Position(lineNumber + 1, 0));
          const lastLineText = activeTextEditor.document.getText(lastLineRange);
          textEditorEdit.replace(lastLineRange, lastLineText);
          return;
        }

        // Try to replace the first character.
        const range = new vscode.Range(new Position(0, 0), new Position(0, 1));
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
          { undoStopBefore: true, undoStopAfter: false });

        await activeTextEditor?.edit(
          (innerEditBuilder: TextEditorEdit) => {
            innerEditBuilder.replace(range, '');
          },
          { undoStopBefore: false, undoStopAfter: true });
      }
    });
  }

  static async closeWatchers(): Promise<void> {
    if (OOXMLViewer.watchers.length) {
      OOXMLViewer.watchers.forEach(w => w.dispose());
      OOXMLViewer.watchers = [];
    }
  }

  private async checkDiff(filePath: string): Promise<void> {
    const ooxmlZip: JSZip = new JSZip();
    const data: Buffer = await readFilePromise(filePath);
    await ooxmlZip.loadAsync(data);
    this.zip = ooxmlZip;
    this.populateOOXMLViewer(this.zip.files);
    this.viewFiles(Object.values(OOXMLViewer.openTextEditors));
  }
}
