import fs from 'fs';
import {join, parse} from 'path';
import JSZip, { JSZipObject } from 'jszip';
import format from 'xml-formatter';
import { OOXMLTreeDataProvider, FileNode } from './ooxml-tree-view-provider';
import vscode, { Uri, TextDocument } from 'vscode';
import mkdirp from 'mkdirp';

/**
 * The OOXML Viewer
 */
export class OOXMLViewer {
    treeDataProvider: OOXMLTreeDataProvider;
    zip: JSZip;

    constructor() {
        this.treeDataProvider = new OOXMLTreeDataProvider();
        this.zip = new JSZip();
    }

    /**
     * Loads the selected OOXML file into the tree view
     * 
     * @param file The OOXML file
     */
    async viewContents(file: vscode.Uri) {
        try {
            this.resetOOXMLViewer();
            let data = await fs.promises.readFile(file.fsPath);
            await this.zip.loadAsync(data);
            this.populateOOXMLViewer(this.zip.files);
        } catch(e) {
            console.error(e);
            vscode.window.showErrorMessage(`Could not load ${file.fsPath}`, e);
        }
    }

    /**
     * Displays the selected file
     * 
     * @param fileNode The selected file node
     */
    async viewFile(fileNode: FileNode) {
        try {
            const file: JSZipObject | null = this.zip.file(fileNode.fullPath);
            const text: string = await file?.async('text') ?? '';
            const formattedXml: string = format(text);
            const root: string = vscode.workspace.rootPath ?? parse(process.cwd()).root;
            const folderPath = join(root, '.ooxml-temp-file-folder-78kIPsmTq5TK');
            const filePath: string = join(folderPath, 'test-file.xml');
            const created: string | void = await mkdirp(folderPath);
            await fs.promises.writeFile(filePath, formattedXml, 'utf8');
            // let xmlDoc = await vscode.workspace.openTextDocument({
            //     content: formattedXml || text
            // });
            const xmlDoc: TextDocument = await vscode.workspace.openTextDocument(Uri.parse("file:///" + filePath));
    
            vscode.window.showTextDocument(xmlDoc);
        } catch(e) {
            console.error(e);
            vscode.window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
        }
    }

    /**
     * Clears the OOXML viewer
     */
    clear() {
        this.resetOOXMLViewer();
    }

    private resetOOXMLViewer() {
        this.zip = new JSZip();
        this.treeDataProvider.rootFileNode = new FileNode();
        this.treeDataProvider.refresh();
    }

    private populateOOXMLViewer(files: {[key: string]: JSZip.JSZipObject}) {
        for (let fileWithPath of Object.keys(files)) {
            // ignore folder files
            if (files[fileWithPath].dir) {
                continue;
            }

            // Build nodes for each file
            let currentFileNode = this.treeDataProvider.rootFileNode;
            for (let fileOrFolderName of fileWithPath.split('/')) {

                // Create node if it does not exist
                let existingFileNode = currentFileNode.children.find(c => c.description === fileOrFolderName);
                if (existingFileNode) {
                    currentFileNode = existingFileNode;
                } else {
                    let newFileNode = new FileNode();
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