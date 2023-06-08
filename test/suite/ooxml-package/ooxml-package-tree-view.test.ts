import { expect } from 'chai';
import { createStubInstance, SinonStubbedInstance } from 'sinon';
import { OOXMLPackageTreeView } from '../../../src/ooxml-package/ooxml-package-tree-view';
import { FileNode, OOXMLTreeDataProvider } from '../../../src/tree-view/ooxml-tree-view-provider';

suite('OOXMLPackageTreeView', function () {
  let treeViewDataProvider: SinonStubbedInstance<OOXMLTreeDataProvider>;
  let treeView: OOXMLPackageTreeView;
  let treeViewRoot: FileNode;

  setup(function () {
    treeViewRoot = new FileNode();
    treeViewDataProvider = createStubInstance(OOXMLTreeDataProvider);
    treeViewDataProvider.rootFileNode = treeViewRoot;

    treeView = new OOXMLPackageTreeView(treeViewDataProvider, 'path/to/package.docx');
  });

  test('getRootNode should return the root tree node', function () {
    const rootNode = treeView.getRootFileNode();

    expect(rootNode.isOOXMLPackage).to.be.true;
    expect(rootNode.nodePath).to.be.eq('path/to/package.docx');
    expect(rootNode.parent).to.be.eq(treeViewRoot);
  });

  test('refresh should call tree view refresh', function () {
    const originalCallCount = treeViewDataProvider.refresh.callCount;
    treeView.refresh();

    expect(treeViewDataProvider.refresh.callCount - originalCallCount).be.eq(1);
  });

  test('reset should remove package root node from the tree view', function () {
    treeView.reset();

    expect(treeViewRoot.children.length).to.be.eq(0);
  });
});
