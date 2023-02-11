import { commands, ExtensionContext, Uri, window } from 'vscode';
import { FileNode } from './ooxml-tree-view-provider';
import { OOXMLViewer } from './ooxml-viewer';

let ooxmlViewer: OOXMLViewer;

export async function activate(context: ExtensionContext): Promise<void> {
  ooxmlViewer = new OOXMLViewer(context);
  await ooxmlViewer.reset();

  context.subscriptions.push(
    window.registerTreeDataProvider('ooxmlViewer', ooxmlViewer.treeDataProvider),
    commands.registerCommand('ooxmlViewer.openOoxmlPackage', async (file: Uri) => ooxmlViewer.openOoxmlPackage(file)),
    commands.registerCommand('ooxmlViewer.removeOoxmlPackage', async (fileNode: FileNode) => fileNode.ooxmlPackage?.reset()),
    commands.registerCommand('ooxmlViewer.viewFile', async (fileNode: FileNode) => fileNode.ooxmlPackage?.viewFile(fileNode.fullPath)),
    commands.registerCommand('ooxmlViewer.clear', () => ooxmlViewer.reset()),
    commands.registerCommand('ooxmlViewer.showDiff', async (fileNode: FileNode) => fileNode.ooxmlPackage?.getDiff(fileNode.fullPath)),
    commands.registerCommand('ooxmlViewer.searchParts', async (fileNode: FileNode) => fileNode.ooxmlPackage?.searchOoxmlParts()),
  );
}

export async function deactivate(): Promise<void> {
  await ooxmlViewer?.reset();
}
