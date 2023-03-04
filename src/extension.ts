import { commands, ExtensionContext, Uri, window } from 'vscode';
import { OOXMLViewer } from './ooxml-viewer';
import { FileNode, OOXMLTreeDataProvider } from './tree-view/ooxml-tree-view-provider';

import packageJson from '../package.json';
const extensionName = packageJson.displayName;

let ooxmlViewer: OOXMLViewer;

export async function activate(context: ExtensionContext): Promise<void> {
  const treeDataProvider = new OOXMLTreeDataProvider();
  const treeView = window.createTreeView('ooxmlViewer', { treeDataProvider: treeDataProvider });
  treeView.title = extensionName;

  ooxmlViewer = new OOXMLViewer(treeDataProvider, context);
  await ooxmlViewer.reset();

  context.subscriptions.push(
    treeView,

    window.registerTreeDataProvider('ooxmlViewer', treeDataProvider),
    commands.registerCommand('ooxmlViewer.openOoxmlPackage', async (file: Uri) => ooxmlViewer.openOOXMLPackage(file.fsPath)),
    commands.registerCommand('ooxmlViewer.removeOoxmlPackage', async (fileNode: FileNode) =>
      ooxmlViewer.removeOOXMLPackage(fileNode.ooxmlPackagePath),
    ),
    commands.registerCommand('ooxmlViewer.viewFile', async (fileNode: FileNode) =>
      ooxmlViewer.viewFile(fileNode.ooxmlPackagePath, fileNode.nodePath),
    ),
    commands.registerCommand('ooxmlViewer.clear', () => ooxmlViewer.reset()),
    commands.registerCommand('ooxmlViewer.showDiff', async (fileNode: FileNode) =>
      ooxmlViewer.getDiff(fileNode.ooxmlPackagePath, fileNode.nodePath),
    ),
    commands.registerCommand('ooxmlViewer.searchParts', async (fileNode: FileNode) =>
      ooxmlViewer.searchOOXMLParts(fileNode.ooxmlPackagePath),
    ),
  );
}

export async function deactivate(): Promise<void> {
  await ooxmlViewer?.reset();
}
