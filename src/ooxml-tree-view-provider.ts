import vscode from 'vscode';

/**
 * OOXML tree data provider
 */
export class OOXMLTreeDataProvider implements vscode.TreeDataProvider<FileNode> {
    private _onDidChangeTreeData: vscode.EventEmitter<FileNode | undefined | null> = new vscode.EventEmitter<FileNode | undefined | null>();

    /**
     * An optional event to signal that an element or root has changed.
     * This will trigger the view to update the changed element/root and its children recursively (if shown).
     * To signal that root has changed, do not pass any argument or pass `undefined` or `null`.
     */
    onDidChangeTreeData?: vscode.Event<FileNode | undefined | null> = this._onDidChangeTreeData.event;

    rootFileNode: FileNode;

    constructor() {
        this.rootFileNode = new FileNode();
    }

    refresh() {
        this._onDidChangeTreeData.fire(undefined);
    }

    /**
     * Get [TreeItem](#TreeItem) representation of the `element`
     *
     * @param element The element for which [TreeItem](#TreeItem) representation is asked for.
     * @return [TreeItem](#TreeItem) representation of the element
     */
    getTreeItem(element: FileNode): vscode.TreeItem | Thenable<vscode.TreeItem> {
        return element;
    }

    /**
     * Get the children of `element` or root if no element is passed.
     *
     * @param element The element from which the provider gets children. Can be `undefined`.
     * @return Children of `element` or root if no element is passed.
     */
    getChildren(element?: FileNode): vscode.ProviderResult<FileNode[]> {
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
    getParent?(element: FileNode): vscode.ProviderResult<FileNode> {
        return element.parent;
    }
}

/**
 * File tree node
 */
export class FileNode implements vscode.TreeItem {
    get description() {
        return this.fileName;
    }

    get collapsibleState()  {
        return this.children.length ? vscode.TreeItemCollapsibleState.Collapsed : vscode.TreeItemCollapsibleState.None;
    }

	get command() {
        if (this.fullPath) {
            return {
                command: 'ooxmlViewer.viewFile',
                title: 'View file',
                tooltip: 'View file tooltip',
                arguments: [this]
            };
        }
    }

    get iconPath() {
        return this.children.length ? vscode.ThemeIcon.Folder : vscode.ThemeIcon.File;
    }

    get tooltip() {
        return this.fullPath;
    }
    
    /**
     * Children of the file node
     */
    children: FileNode[] = [];

    /**
     * Full path of the file
     */
    fullPath: string = '';

    /**
     * Name of the file
     */
    fileName: string = '';

    /**
     * Parent file node
     */
    parent: FileNode | undefined;
}