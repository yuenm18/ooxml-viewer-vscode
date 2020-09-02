import { exec } from 'child_process';
import { existsSync, readFile, writeFile } from 'fs';
import JSZip, { JSZipObject } from 'jszip';
import mkdirp from 'mkdirp';
import { dirname, join, parse } from 'path';
import rimraf from 'rimraf';
import { promisify } from 'util';
import vscode, { TextDocument, Uri } from 'vscode';
import format from 'xml-formatter';
import { FileNode, OOXMLTreeDataProvider } from './ooxml-tree-view-provider';
const execPromise = promisify(exec);
const readFilePromise = promisify(readFile);
const writeFilePromise = promisify(writeFile);
const rimrafPromise = promisify(rimraf);

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
  treeDataProvider: OOXMLTreeDataProvider;
  zip: JSZip;
  static p = parse(process.cwd());
  static fileCachePath: string = join(vscode.workspace.rootPath == undefined ? OOXMLViewer.p.root : vscode.workspace.rootPath, '.ooxml-temp-file-folder-78kIPsmTq5TK');
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
      this.resetOOXMLViewer();
      const data = await readFilePromise(file.fsPath);
      await this.zip.loadAsync(data);
      this.populateOOXMLViewer(this.zip.files);
      // TODO: Use this watch to update the ooxml file when the file is changed from outside vscode. i.e. in PowerPoint
      // watch(file.fsPath, { encoding: 'buffer' }, (eventType: string, filename: Buffer): void => {
      //     vscode.window.showWarningMessage('I am a warning', { modal: true }, ...['Save', 'Discard']);
      //     if (filename) {
      //         console.log(filename);
      //         // Prints: <Buffer ...>
      //     }
      // });
    } catch (e) {
      console.error(e);
      vscode.window.showErrorMessage(`Could not load ${file.fsPath}`, e);
    }
  }

  /**
   * Displays the selected file
   *
   * @param fileNode The selected file node
   */
  async viewFile(fileNode: FileNode): Promise<void> {
    try {
      const file: JSZipObject | null = this.zip.file(fileNode.fullPath);
      const text: string = await file?.async('text') ?? '';
      const formattedXml: string = format(text);
      const folderPath = join(OOXMLViewer.fileCachePath, dirname(fileNode.fullPath));
      const filePath: string = join(folderPath, fileNode.fileName);
      await OOXMLViewer.mkdirp(folderPath);
      // On Windows hide the folder
      if (process.platform.startsWith('win')) {
        const { stderr } = await OOXMLViewer.execPromise('attrib +h ' + OOXMLViewer.fileCachePath);
        if (stderr) {
          throw new Error(stderr);
        }
      }
      await OOXMLViewer.writeFilePromise(filePath, formattedXml, 'utf8');
      const xmlDoc: TextDocument = await vscode.workspace.openTextDocument(Uri.parse('file:///' + filePath));

      vscode.window.showTextDocument(xmlDoc);
    } catch (e) {
      console.error(e);
      vscode.window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
    }
  }

  /**
   * Clears the OOXML viewer
   */
  clear(): Promise<void> {
    return this.resetOOXMLViewer();
  }
  private static closeEditors(tds: TextDocument[]): void {
    if (tds.length) {
      const td: TextDocument | undefined = tds.pop();
      if (td) {
        vscode.window.showTextDocument(td, { preview: true, preserveFocus: false })
          .then(() => {
            vscode.commands.executeCommand('workbench.action.closeActiveEditor');
          })
          .then(() => {
            OOXMLViewer.closeEditors(tds);
          });
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
      OOXMLViewer.closeEditors(vscode.workspace.textDocuments.filter(t => t.fileName.toLowerCase().includes(OOXMLViewer.fileCachePath.toLowerCase())));
    } catch (err) {
      console.error(err);
      vscode.window.showErrorMessage('Could not remove ooxml file viewer cache');
    }
  }

  private populateOOXMLViewer(files: { [key: string]: JSZip.JSZipObject; }) {
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
        }
      }

      currentFileNode.fullPath = fileWithPath;
    }

    this.treeDataProvider.refresh();
  }
}
