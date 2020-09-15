import rimraf from 'rimraf';
import { promisify } from 'util';
import * as vscode from 'vscode';
import { FileNode } from './ooxml-tree-view-provider';
import { OOXMLViewer } from './ooxml-viewer';

const rimrafPromise = promisify(rimraf);

export function activate(context: vscode.ExtensionContext): void {
  const ooxmlViewer = new OOXMLViewer();
  context.subscriptions.push(vscode.window.registerTreeDataProvider('ooxmlViewer', ooxmlViewer.treeDataProvider));
  context.subscriptions.push(vscode.commands.registerCommand('ooxmlViewer.viewContents', async (file: vscode.Uri) => ooxmlViewer.viewContents(file)));
  context.subscriptions.push(vscode.commands.registerCommand('ooxmlViewer.viewFile', async (fileNode: FileNode) => ooxmlViewer.viewFile(fileNode)));
  context.subscriptions.push(vscode.commands.registerCommand('ooxmlViewer.clear', () => ooxmlViewer.clear()));
}

export async function deactivate(): Promise<void> {
  if (OOXMLViewer.watchers.length) {
    OOXMLViewer.watchers.forEach(w => w.close());
  }
  return rimrafPromise(OOXMLViewer.fileCachePath);
}
