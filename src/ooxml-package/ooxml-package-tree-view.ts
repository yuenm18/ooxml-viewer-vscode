import { FileNode, OOXMLTreeDataProvider } from '../tree-view/ooxml-tree-view-provider';

/**
 * The ooxml package tree view. Represents the part of the tree view that contains the ooxml file.
 */
export class OOXMLPackageTreeView {
  private rootFileNode: FileNode;

  /**
   * Creates a new instance of a root file node.
   *
   * @constructor
   * @param {OOXMLTreeDataProvider} treeDataProvider The tree data provider.
   * @param {string} ooxmlPackagePath The path to the ooxml file.
   */
  constructor(private treeDataProvider: OOXMLTreeDataProvider, private ooxmlPackagePath: string) {
    this.rootFileNode = this.createRootNode();
  }

  /**
   * Return the root file node of the ooxml package.
   *
   * @returns {FileNode} Returns the root file node of the ooxml package.
   */
  getRootFileNode(): FileNode {
    return this.rootFileNode;
  }

  /**
   * Refresh the tree view.
   */
  refresh(): void {
    this.treeDataProvider.refresh();
  }

  /**
   * Resets the ooxml package tree and removes it from the tree view.
   */
  reset(): void {
    const nodeIndex = this.treeDataProvider.rootFileNode.children.indexOf(this.rootFileNode);
    if (nodeIndex !== -1) {
      this.treeDataProvider.rootFileNode.children.splice(nodeIndex, 1);
      this.treeDataProvider.refresh();
    }
  }

  private createRootNode(): FileNode {
    const rootFileNode = FileNode.create(this.ooxmlPackagePath, this.treeDataProvider.rootFileNode, this.ooxmlPackagePath);
    rootFileNode.isOOXMLPackage = true;
    this.treeDataProvider.refresh();

    return rootFileNode;
  }
}
