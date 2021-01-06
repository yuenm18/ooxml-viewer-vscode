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
  getParent?(element: FileNode): ProviderResult<FileNode> {
    return element.parent;
  }
}

/**
 * File tree node
 */
export class FileNode implements TreeItem {
  private _status: 'created' | 'deleted' | 'modified' | 'unchanged' = 'unchanged'

  get description(): string {
    return this.fileName;
  }

  get collapsibleState(): TreeItemCollapsibleState | undefined {
    return this.children.length ? TreeItemCollapsibleState.Expanded : TreeItemCollapsibleState.None;
  }

  get command(): Command | undefined {
    if (this.fullPath) {
      return {
        command: this.children.length ? '' : 'ooxmlViewer.viewFile',
        title: 'View file',
        tooltip: 'View file tooltip',
        arguments: [this],
      };
    }
  }

  get iconPath(): ThemeIcon | Uri | { light: Uri; dark: Uri } {
    switch (this._status) {
      case 'created': return Uri.file(join(__filename, '..', '..', 'resources', 'icons', 'asterisk.green.svg'));
      case 'deleted': return Uri.file(join(__filename, '..', '..', 'resources', 'icons', 'asterisk.red.svg'));
      case 'modified': return Uri.file(join(__filename, '..', '..', 'resources', 'icons', 'asterisk.yellow.svg'));
      default: return this.children.length ? ThemeIcon.Folder : ThemeIcon.File;
    }
  }

  get tooltip(): string {
    return this.fullPath;
  }

  /**
   * Children of the file node
   */
  children: FileNode[] = [];

  /**
   * Full path of the file
   */
  fullPath = '';

  /**
   * Name of the file
   */
  fileName = '';

  /**
   * Parent file node
   */
  parent: FileNode | undefined;

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
