import * as vscode from 'vscode';
import { FileNode } from './ooxml-tree-view-provider';
import { OOXMLViewer } from './ooxml-viewer';

export function activate(context: vscode.ExtensionContext) {
	const ooxmlViewer = new OOXMLViewer();
	context.subscriptions.push(vscode.window.registerTreeDataProvider('ooxmlViewer', ooxmlViewer.treeDataProvider));
	context.subscriptions.push(vscode.commands.registerCommand('ooxmlViewer.viewContents', async (file: vscode.Uri) => ooxmlViewer.viewContents(file)));
	context.subscriptions.push(vscode.commands.registerCommand('ooxmlViewer.viewFile', async (fileNode: FileNode) => ooxmlViewer.viewFile(fileNode)));
	context.subscriptions.push(vscode.commands.registerCommand('ooxmlViewer.clear', () => ooxmlViewer.clear()));
}

export function deactivate() {}
