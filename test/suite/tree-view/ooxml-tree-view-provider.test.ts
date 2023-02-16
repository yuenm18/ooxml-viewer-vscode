import { expect } from 'chai';
import { SinonStub, spy } from 'sinon';
import { EventEmitter, ThemeIcon, TreeItemCollapsibleState, Uri } from 'vscode';
import { FileNode, OOXMLTreeDataProvider } from '../../../src/tree-view/ooxml-tree-view-provider';

suite('OOXMLViewer Tree View Provider', function () {
  const stubs: SinonStub[] = [];
  let treeViewProvider: OOXMLTreeDataProvider;

  setup(function () {
    treeViewProvider = new OOXMLTreeDataProvider();
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('should have a rootFileNode that is an instance of FileNode', function () {
    expect(treeViewProvider.rootFileNode).to.be.an.instanceof(FileNode);
  });

  test("should return root file node's children if no node is passed in when getChildren is called", function () {
    expect(treeViewProvider.getChildren()).to.be.equal(treeViewProvider.rootFileNode.children);
  });

  test("should return file node's children if node is passed in when getChildren is called", function () {
    const fileNode = new FileNode();
    fileNode.children = [new FileNode(), new FileNode()];

    expect(treeViewProvider.getChildren(fileNode)).to.be.equal(fileNode.children);
  });

  test("should return file node's parent when getParent is called", function () {
    const fileNode = new FileNode();
    fileNode.parent = new FileNode();

    expect(treeViewProvider.getParent(fileNode)).to.be.equal(fileNode.parent);
  });

  test('should return file node when getTreeItem is called', function () {
    const fileNode = new FileNode();

    expect(treeViewProvider.getTreeItem(fileNode)).to.be.equal(fileNode);
  });

  test('should call _onDidChangeTreeData.fire when refresh is called', function () {
    const fireSpy = spy();
    treeViewProvider['_onDidChangeTreeData'] = {
      fire: fireSpy,
    } as unknown as EventEmitter<FileNode | null | undefined>;
    treeViewProvider.refresh();
    expect(fireSpy.calledWith(undefined)).to.be.true;
  });
});

suite('OOXMLViewer File Node', function () {
  let fileNode: FileNode;

  setup(function () {
    fileNode = new FileNode();
    fileNode.fullPath = 'tmp/file.docx';
    fileNode.fileName = 'file.docx';
  });

  test('should have file icon if fileNode has no children', function () {
    fileNode.children = [];

    expect(fileNode.iconPath).to.be.equal(ThemeIcon.File);
  });

  test('should have folder icon if fileNode has children', function () {
    fileNode.children = [new FileNode()];

    expect(fileNode.iconPath).to.be.equal(ThemeIcon.Folder);
  });

  test('should have context value of "package" if fileNode is an ooxml package', function () {
    fileNode.isOOXMLPackage = true;

    expect(fileNode.contextValue).to.be.equal('package');
  });

  test('should have context value of "file" if fileNode has no children', function () {
    fileNode.children = [];

    expect(fileNode.contextValue).to.be.equal('file');
  });

  test('should context value of "folder" if fileNode has children', function () {
    fileNode.children = [new FileNode()];

    expect(fileNode.contextValue).to.be.equal('folder');
  });

  test('should have asterisk.green.svg if fileNode is created', function () {
    fileNode.setCreated();

    expect((fileNode.iconPath as Uri)?.fsPath).to.contain('asterisk.green.svg');
  });

  test('should have asterisk.red.svg if fileNode is deleted', function () {
    fileNode.setDeleted();

    expect((fileNode.iconPath as Uri)?.fsPath).to.contain('asterisk.red.svg');
  });

  test('should have asterisk.yellow.svg if fileNode is modified', function () {
    fileNode.setModified();

    expect((fileNode.iconPath as Uri)?.fsPath).to.contain('asterisk.yellow.svg');
  });

  test('should be unset if fileNode is set to unchanged()', function () {
    fileNode.setModified();
    fileNode.setUnchanged();

    expect(fileNode.iconPath).to.be.instanceOf(ThemeIcon);
  });

  test('should have isDeleted() = true if fileNode is deleted', function () {
    fileNode.setDeleted();

    expect(fileNode.isDeleted()).to.be.true;
  });

  test('should have TreeItemCollapsibleState.None if fileNode has no children', function () {
    fileNode.children = [];

    expect(fileNode.collapsibleState).to.be.equal(TreeItemCollapsibleState.None);
  });

  test('should have TreeItemCollapsibleState.Expanded if fileNode has children', function () {
    fileNode.children = [new FileNode()];

    expect(fileNode.collapsibleState).to.be.equal(TreeItemCollapsibleState.Expanded);
  });

  test('should not return command if fileNode has children', function () {
    fileNode.children = [new FileNode()];

    expect(fileNode.command).to.be.undefined;
  });

  test('should return viewFile command if fileNode has no children', function () {
    fileNode.fullPath = '/file.docx';
    fileNode.children = [];

    expect(fileNode.command?.command).to.be.equal('ooxmlViewer.viewFile');
    expect(fileNode.command?.arguments).to.have.members([fileNode]);
  });

  test('should have a tooltip as the full path', function () {
    expect(fileNode.tooltip).to.be.equal(fileNode.fullPath);
  });

  test('should have the file name as a description', function () {
    expect(fileNode.description).to.be.equal(fileNode.fileName);
  });
});
