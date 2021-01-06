import { commands, ExtensionContext, Uri, window } from 'vscode';
import { FileNode } from './ooxml-tree-view-provider';
import { OOXMLViewer } from './ooxml-viewer';

let ooxmlViewer: OOXMLViewer;

export function activate(context: ExtensionContext): void {
  ooxmlViewer = new OOXMLViewer(context);
  
  context.subscriptions.push(window.registerTreeDataProvider('ooxmlViewer', ooxmlViewer.treeDataProvider));
  context.subscriptions.push(
    commands.registerCommand('ooxmlViewer.openOoxmlPackage', async (file: Uri) => ooxmlViewer.openOoxmlPackage(file)),
  );
  context.subscriptions.push(
    commands.registerCommand('ooxmlViewer.viewFile', async (fileNode: FileNode) => ooxmlViewer.viewFile(fileNode)),
  );
  context.subscriptions.push(commands.registerCommand('ooxmlViewer.clear', () => ooxmlViewer.clear()));
  context.subscriptions.push(commands.registerCommand('ooxmlViewer.showDiff', async (file: FileNode) => ooxmlViewer.getDiff(file)));
}

export async function deactivate(): Promise<void> {
  ooxmlViewer?.disposeWatchers();
}
