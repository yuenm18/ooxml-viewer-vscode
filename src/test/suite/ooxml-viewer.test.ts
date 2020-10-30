import { expect } from 'chai';
import { dirname, join } from 'path';
import { SinonStub, stub } from 'sinon';
import { ImportMock } from 'ts-mock-imports';
import { commands, ExtensionContext, Uri, workspace } from 'vscode';
import { FileNode, OOXMLTreeDataProvider } from '../../ooxml-tree-view-provider';
import { OOXMLViewer } from '../../ooxml-viewer';

suite('OOXMLViewer', function () {
  // this.timeout(10000);
  let ooxmlViewer: OOXMLViewer;
  const stubs: SinonStub[] = [];
  const testFilePath = join(__dirname, '..', '..', '..', 'test-data', 'Test.pptx');
  let context: { [key: string]: (path?: string | undefined) => string };

  setup(function () {
    context = {
      asAbsolutePath: (path?: string) => {
        return 'tacocat';
      },
    };
    ooxmlViewer = new OOXMLViewer((context as unknown) as ExtensionContext);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
  });

  test('It should have an instance of OOXMLTreeDataProvider', function (done) {
    expect(ooxmlViewer.treeDataProvider).to.be.instanceOf(OOXMLTreeDataProvider);
    done();
  });
  test('It should populate the sidebar tree with the contents of an ooxml file', async function () {
    const mkdirMock = ImportMock.mockFunction(workspace.fs, 'createDirectory').returns(Promise.resolve());
    const createFileMock = stub(OOXMLViewer.prototype, <never>'_createFile').returns(Promise.resolve());

    stubs.push(stub(OOXMLViewer, <never>'_fileHasBeenChangedFromOutside').returns(Promise.resolve(false)), mkdirMock, createFileMock);
    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(0);
    await ooxmlViewer.viewContents(Uri.file(testFilePath));
    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(4);
    expect(createFileMock.callCount).to.eq(225);
  });
  test('viewFile should open a text editor when called with the path to an xml file', function (done) {
    const commandsMock = ImportMock.mockFunction(commands, 'executeCommand');
    stubs.push(
      stub(OOXMLViewer.prototype, <never>'_createFile').callsFake(() => Promise.resolve()),
      stub(workspace.fs, 'readFile').callsFake((uri: Uri) => {
        console.log('uri', uri);
        const enc = new TextEncoder();
        return Promise.resolve(
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
              "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>",
          ),
        );
      }),
      stub(OOXMLViewer, <never>'_fileHasBeenChangedFromOutside').returns(Promise.resolve(false)),
      commandsMock,
    );
    const node = new FileNode();
    ooxmlViewer.viewFile(node).then(() => {
      const folderPath = join(OOXMLViewer.fileCachePath, dirname(node.fullPath));
      const filePath: string = join(folderPath, node.fileName);
      expect(OOXMLViewer.openTextEditors[filePath]).to.eq(node);
      expect(commandsMock.calledWith('vscode.open')).to.be.true;
      done();
    });
  });
  suiteTeardown(function () {
    stubs.forEach(s => s.restore());
  });
});
