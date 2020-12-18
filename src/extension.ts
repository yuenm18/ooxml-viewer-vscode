import { commands, ExtensionContext, Uri, window } from 'vscode';
import { FileNode } from './ooxml-tree-view-provider';
import { OOXMLViewer } from './ooxml-viewer';

export function activate(context: ExtensionContext): void {
  const ooxmlViewer = new OOXMLViewer(context);
  ooxmlViewer.closeEditors();
  context.subscriptions.push(window.registerTreeDataProvider('ooxmlViewer', ooxmlViewer.treeDataProvider));
  context.subscriptions.push(commands.registerCommand('ooxmlViewer.viewContents', async (file: Uri) => ooxmlViewer.viewContents(file)));
  context.subscriptions.push(
    commands.registerCommand('ooxmlViewer.viewFile', async (fileNode: FileNode) => ooxmlViewer.viewFile(fileNode)),
  );
  context.subscriptions.push(commands.registerCommand('ooxmlViewer.clear', () => ooxmlViewer.clear()));
  context.subscriptions.push(commands.registerCommand('ooxmlViewer.showDiff', async (file: FileNode) => ooxmlViewer.getDiff(file)));
}

export async function deactivate(): Promise<void> {
  OOXMLViewer.closeWatchers();
}
