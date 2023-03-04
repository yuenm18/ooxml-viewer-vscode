import { join } from 'path';
import { Command, Event, EventEmitter, ProviderResult, ThemeIcon, TreeDataProvider, TreeItem, TreeItemCollapsibleState, Uri } from 'vscode';

/**
 * OOXML tree data provider
 */
export class OOXMLTreeDataProvider implements TreeDataProvider<FileNode> {
  private _onDidChangeTreeData: EventEmitter<FileNode | undefined | null> = new EventEmitter<FileNode | undefined | null>();

  /**
   * An optional event to signal that an element or root has changed.
   * This will trigger the view to update the changed element/root and its children recursively (if shown).
   * To signal that root has changed, do not pass any argument or pass `undefined` or `null`.
   */
  onDidChangeTreeData?: Event<FileNode | undefined | null> = this._onDidChangeTreeData.event;

  rootFileNode: FileNode;

  constructor() {
    this.rootFileNode = new FileNode();
  }

  refresh(): void {
    this._onDidChangeTreeData.fire(undefined);
  }

  /**
   * Get [TreeItem](#TreeItem) representation of the `element`
   *
   * @param element The element for which [TreeItem](#TreeItem) representation is asked for.
   * @return [TreeItem](#TreeItem) representation of the element
   */
  getTreeItem(element: FileNode): TreeItem | Thenable<TreeItem> {
    return element;
  }

  /**
   * Get the children of `element` or root if no element is passed.
   *
   * @param element The element from which the provider gets children. Can be `undefined`.
   * @return Children of `element` or root if no element is passed.
   */
  getChildren(element?: FileNode): ProviderResult<FileNode[]> {
    return element ? element.children : this.rootFileNode.children;
  }

  /**
   * Optional method to return the parent of `element`.
   * Return `null` or `undefined` if `element` is a child of root.
   *
   * **NOTE:** This method should be implemented in order to access [reveal](#TreeView.reveal) API.
   *
   * @param element The element for which the parent has to be returned.
   * @return Parent of `element`.
   */
  getParent(element: FileNode): ProviderResult<FileNode> {
    return element.parent;
  }
}

/**
 * File tree node
 */
export class FileNode implements TreeItem {
  /**
   * Creates a file node.
   *
   * @param  {string} nodePath The path of the FileNode.
   * @param  {FileNode} parentFileNode The parent FileNode.
   * @returns {FileNode}
   */
  public static create(nodePath: string, parentFileNode: FileNode, ooxmlPackagePath: string): FileNode {
    const fileNode = new FileNode();
    fileNode.nodePath = nodePath;
    fileNode.parent = parentFileNode;
    fileNode.ooxmlPackagePath = ooxmlPackagePath;
    parentFileNode.children.push(fileNode);
    return fileNode;
  }

  private _status: 'created' | 'deleted' | 'modified' | 'unchanged' = 'unchanged';

  get collapsibleState(): TreeItemCollapsibleState | undefined {
    return this.contextValue === FileNodeType.File ? TreeItemCollapsibleState.None : TreeItemCollapsibleState.Expanded;
  }

  get command(): Command | undefined {
    if (this.nodePath && this.contextValue === FileNodeType.File) {
      return {
        command: 'ooxmlViewer.viewFile',
        title: 'View file',
        tooltip: 'View file',
        arguments: [this],
      };
    }
  }

  get contextValue(): FileNodeType | undefined {
    return this.isOOXMLPackage ? FileNodeType.Package : this.children.length ? FileNodeType.Folder : FileNodeType.File;
  }

  get resourceUri(): Uri {
    return Uri.file(this.nodePath);
  }

  get tooltip(): string {
    return this.nodePath;
  }

  get iconPath(): ThemeIcon | Uri | { light: Uri; dark: Uri } {
    if (this.isOOXMLPackage) {
      return new ThemeIcon('package');
    }

    switch (this._status) {
      case 'created':
        return Uri.file(join(__filename, '..', '..', 'resources', 'icons', 'asterisk.green.svg'));
      case 'deleted':
        return Uri.file(join(__filename, '..', '..', 'resources', 'icons', 'asterisk.red.svg'));
      case 'modified':
        return Uri.file(join(__filename, '..', '..', 'resources', 'icons', 'asterisk.yellow.svg'));
      default:
        return this.contextValue === FileNodeType.File ? ThemeIcon.File : ThemeIcon.Folder;
    }
  }

  /**
   * Children of the file node
   */
  children: FileNode[] = [];

  /**
   * Full path of the file or folder
   */
  nodePath = '';

  /**
   * Parent file node
   */
  parent: FileNode | undefined;

  /**
   * The ooxml package the file node is part of
   */
  ooxmlPackagePath = '';

  /**
   * Whether or not the file node is a ooxml package (root level)
   */
  isOOXMLPackage = false;

  /**
   * Gets whether or not the file node has a status of deleted.
   *
   * @returns True if status is deleted, false otherwise.
   */
  isDeleted(): boolean {
    return this._status === 'deleted';
  }

  /**
   * Sets the status of the file node to created.
   */
  setCreated(): void {
    this._status = 'created';
  }

  /**
   * Sets the status of the file node to deleted.
   */
  setDeleted(): void {
    this._status = 'deleted';
  }

  /**
   * Sets the status of the file node to modified.
   */
  setModified(): void {
    this._status = 'modified';
  }

  /**
   * Sets the status of the file node to unchanged.
   */
  setUnchanged(): void {
    this._status = 'unchanged';
  }
}

export enum FileNodeType {
  Package = 'package',
  Folder = 'folder',
  File = 'file',
}
