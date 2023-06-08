import { commands, ExtensionContext, Uri, window } from 'vscode';
import { OOXMLViewer } from './ooxml-viewer';
import { FileNode, OOXMLTreeDataProvider } from './tree-view/ooxml-tree-view-provider';

import packageJson from '../package.json';
import { getExtensionSettings } from './ooxml-extension-settings';
import logger from './utilities/logger';

const extensionName = packageJson.displayName;

let ooxmlViewer: OOXMLViewer;

export async function activate(context: ExtensionContext): Promise<void> {
  const treeDataProvider = new OOXMLTreeDataProvider();
  const treeView = window.createTreeView('ooxmlViewer', { treeDataProvider: treeDataProvider });
  treeView.title = extensionName;

  const settings = getExtensionSettings();
  logger.info(`Starting '${extensionName}': ${JSON.stringify(settings, null, 4)}`);

  ooxmlViewer = new OOXMLViewer(treeDataProvider, settings, context);
  await ooxmlViewer.reset();

  context.subscriptions.push(
    treeView,

    window.registerTreeDataProvider('ooxmlViewer', treeDataProvider),
    commands.registerCommand('ooxmlViewer.openOoxmlPackage', (file: Uri) => ooxmlViewer.openOOXMLPackage(file.fsPath)),
    commands.registerCommand('ooxmlViewer.removeOoxmlPackage', (fileNode: FileNode) =>
      ooxmlViewer.removeOOXMLPackage(fileNode.ooxmlPackagePath),
    ),
    commands.registerCommand('ooxmlViewer.viewFile', (fileNode: FileNode) =>
      ooxmlViewer.viewFile(fileNode.ooxmlPackagePath, fileNode.nodePath),
    ),
    commands.registerCommand('ooxmlViewer.clear', () => ooxmlViewer.reset()),
    commands.registerCommand('ooxmlViewer.showDiff', (fileNode: FileNode) =>
      ooxmlViewer.getDiff(fileNode.ooxmlPackagePath, fileNode.nodePath),
    ),
    commands.registerCommand('ooxmlViewer.searchParts', (fileNode: FileNode) => ooxmlViewer.searchOOXMLParts(fileNode.ooxmlPackagePath)),
  );
}

export async function deactivate(): Promise<void> {
  logger.info(`Deactivating '${extensionName}'`);
  await ooxmlViewer?.reset();
}
