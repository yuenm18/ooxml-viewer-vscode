import * as fs from 'fs';
import * as JSZip from 'jszip';
import { OOXMLTreeDataProvider, FileNode } from './ooxml-tree-view-provider';
import * as vscode from 'vscode';

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
            let file = this.zip.file(fileNode.fullPath);
            let text = await file.async("text");
            let xmlDoc = await vscode.workspace.openTextDocument({
                content: text
            });
    
            vscode.window.showTextDocument(xmlDoc);
        } catch(e) {
            console.error(e);
            vscode.window.showErrorMessage(`Could not load ${fileNode.fullPath}`);
        }
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