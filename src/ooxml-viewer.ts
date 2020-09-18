import { exec } from 'child_process';
import { existsSync, FSWatcher, readFile, stat, Stats, watch, writeFile } from 'fs';
import JSZip, { JSZipObject } from 'jszip';
import mkdirp from 'mkdirp';
import { dirname, format, join, parse } from 'path';
import rimraf from 'rimraf';
import { promisify } from 'util';
import vscode, { Position, TextDocument, TextEditorEdit, Uri } from 'vscode';
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
  static watchers: FSWatcher[] = [];
  static watchActions: { [key: string]: number; } = {};
  static cacheFolderName = '.ooxml-temp-file-folder-78kIPsmTq5TK';
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
      await this.resetOOXMLViewer();
      const data = await readFilePromise(file.fsPath);
      await this.zip.loadAsync(data);
      await this.populateOOXMLViewer(this.zip.files);
      await OOXMLViewer.mkdirp(OOXMLViewer.fileCachePath);
      // TODO: Use this watch to update the ooxml file when the file is changed from outside vscode. e.g. in PowerPoint
      OOXMLViewer.watchers.push(watch(file.fsPath, { encoding: 'buffer' }, async (eventType: string, filename: Buffer): Promise<void> => {
        this.checkDiff(file.fsPath);
      }));

      OOXMLViewer.watchers.push(watch(OOXMLViewer.fileCachePath, { encoding: 'buffer', recursive: true },
        async (eventType: string, fileNameBuffer: Buffer): Promise<void> => {
          try {
            const name: string | undefined = fileNameBuffer == undefined ? undefined : fileNameBuffer.toString('utf-8');
            const filePath: string | undefined = name ? join(OOXMLViewer.fileCachePath, name) : undefined;
            let prevFilePath = '';
            if (filePath) {
              const { dir, base } = parse(filePath);
              prevFilePath = format({ dir, base: `prev.${base}` });
            }
            if (name && filePath && existsSync(filePath) && prevFilePath && existsSync(prevFilePath)) {
              const stats: Stats = await statPromise(filePath);
              const time = stats.mtime.getTime();
              if (!stats.isDirectory() && eventType === 'change' && OOXMLViewer.watchActions[name] !== time) {
                OOXMLViewer.watchActions[name] = time;
                const data: Buffer = await readFilePromise(filePath);
                const prevData: Buffer = await readFilePromise(prevFilePath);
                if (!data.equals(prevData)) {
                  const normalizedPath: string = name.replace(/\\/g, '/');
                  const zipFile = await this.zip.file(normalizedPath, data, { binary: true }).generateAsync({ type: 'nodebuffer' });
                  await writeFilePromise(file.fsPath, zipFile);
                  await writeFilePromise(prevFilePath, data);
                }
              }
            }
          } catch (err) {
            if (err && err.code === 'EBUSY') {
              vscode.window.showWarningMessage(
                `File not saved.\n${file.fsPath} is open in another program.\nClose that program before making any changes.`,
                { modal: true },
              );
              OOXMLViewer.makeDirty();
            }
          }
        }));
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
      // OOXMLViewer.openTextEditors[fileNode.fullPath] = fileNode;
      const folderPath = join(OOXMLViewer.fileCachePath, dirname(fileNode.fullPath));
      const filePath: string = join(folderPath, fileNode.fileName);
      await this.createFile(fileNode.fullPath, fileNode.fileName);
      const xmlDoc: TextDocument = await vscode.workspace.openTextDocument(Uri.parse('file:///' + filePath));

      await vscode.window.showTextDocument(xmlDoc);
    } catch (e) {
      console.error(e);
      vscode.window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
    }
  }

  private async viewFiles(fileNodes: FileNode[]): Promise<void> {
    if (fileNodes.length) {
      const node = fileNodes.pop();
      if (node) {
        await this.viewFile(node);
        await OOXMLViewer.makeDirty();
        // const saved: boolean | undefined = await vscode.window.activeTextEditor?.document.save();
        // console.log('saved', saved);
        setTimeout(() => {
          this.viewFiles(fileNodes);
        }, 10);
      }
    } else {
      // vscode.commands.executeCommand('workbench.action.files.saveAll');
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
          currentFileNode.children.push(newFileNode);
          currentFileNode = newFileNode;
          this.createFile(newFileNode.fullPath, `prev.${newFileNode.fileName}`);
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
      }
    } catch (err) {
      console.error(err);
    }
  }

  private static async makeDirty(): Promise<void> {
    vscode.window.activeTextEditor?.edit(async (textEditorEdit: TextEditorEdit) => {
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
      OOXMLViewer.watchers.forEach(w => w.close());
      OOXMLViewer.watchers = [];
    }
  }

  private async checkDiff(filePath: string): Promise<void> {
    const ooxmlZip: JSZip = new JSZip();
    const data: Buffer = await readFilePromise(filePath);
    await ooxmlZip.loadAsync(data);
    this.zip = ooxmlZip;
    this.populateOOXMLViewer(ooxmlZip.files);
    OOXMLViewer.closeEditors();
  }
}
