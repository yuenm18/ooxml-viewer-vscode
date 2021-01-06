import { expect } from 'chai';
import { SinonStub, spy } from 'sinon';
import { EventEmitter } from 'vscode';
import { FileNode, OOXMLTreeDataProvider } from './../../src/ooxml-tree-view-provider';

suite('OOXMLViewer Tree View Provider', function () {
  const stubs: SinonStub[] = [];
  let treeViewProvider: OOXMLTreeDataProvider;

  suiteSetup(function () {
    treeViewProvider = new OOXMLTreeDataProvider();
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('It should have a rootFileNode that is an instand of FileNode', function (done) {
    expect(treeViewProvider.rootFileNode).to.be.an.instanceof(FileNode);
    done();
  });

  test('It should call _onDidChangeTreeData.fire when refresh is called', function () {
    const fireSpy = spy();
    treeViewProvider['_onDidChangeTreeData'] = ({
      fire: fireSpy,
    } as unknown) as EventEmitter<FileNode | null | undefined>;
    treeViewProvider.refresh();
    expect(fireSpy.calledWith(undefined)).to.be.true;
  });
});
