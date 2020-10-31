import { expect } from 'chai';
import child_process from 'child_process';
import JSZip from 'jszip';
import { dirname, join } from 'path';
import { SinonStub, stub } from 'sinon';
import { commands, ExtensionContext, Uri, workspace } from 'vscode';
import formatXml from 'xml-formatter';
import { FileNode, OOXMLTreeDataProvider } from '../../ooxml-tree-view-provider';
import { OOXMLViewer } from '../../ooxml-viewer';

suite('OOXMLViewer', async function () {
  this.timeout(10000);
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
    const mkdirMock = stub(workspace.fs, 'createDirectory').returns(Promise.resolve());
    const createFileMock = stub(OOXMLViewer.prototype, <never>'_createFile').returns(Promise.resolve());

    stubs.push(stub(OOXMLViewer, <never>'_fileHasBeenChangedFromOutside').returns(Promise.resolve(false)), mkdirMock, createFileMock);
    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(0);
    await ooxmlViewer.viewContents(Uri.file(testFilePath));
    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(4);
    expect(createFileMock.callCount).to.eq(225);
  });
  test('viewFile should open a text editor when called with the path to an xml file', async function () {
    const commandsStub = stub(commands, 'executeCommand');
    const createDirectoryStub = stub(workspace.fs, 'createDirectory').callsFake(uri => {
      expect(uri).to.be.instanceof(Uri);
      return Promise.resolve();
    });
    const spawnStub = stub(child_process, 'spawn').returns({ stderr: null } as child_process.ChildProcess);
    const enc = new TextEncoder();
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
      "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>";
    const coded = enc.encode(xml);
    const readFileStub = stub(workspace.fs, 'readFile').callsFake((uri: Uri) => {
      expect(uri).to.be.instanceof(Uri);
      return Promise.resolve(coded);
    });
    const writeFileStub = stub(workspace.fs, 'writeFile').callsFake((uri: Uri, u8a: Uint8Array) => {
      expect(uri).to.be.instanceof(Uri);
      const sent: Buffer = Buffer.from(enc.encode(formatXml(xml)));
      const received: Buffer = Buffer.from(u8a);
      expect(sent.equals(received)).to.be.true;
      return Promise.resolve();
    });
    const zipStub = stub(ooxmlViewer.zip, 'file').returns(({
      async() {
        return Promise.resolve(
          '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
            "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>",
        );
      },
    } as never) as JSZip);
    stubs.push(commandsStub, createDirectoryStub, spawnStub, readFileStub, writeFileStub, zipStub);
    const node = new FileNode();
    await ooxmlViewer.viewFile(node);
    const folderPath = join(OOXMLViewer.fileCachePath, dirname(node.fullPath));
    const filePath: string = join(folderPath, node.fileName);
    expect(OOXMLViewer.openTextEditors[filePath]).to.eq(node);
    expect(commandsStub.calledWith('vscode.open')).to.be.true;
    if (process && (process.platform === 'win32' || (process.env && process.env.OSTYPE && /^(msys|cygwin)$/.test(process.env.OSTYPE)))) {
      expect(spawnStub.calledWith('attrib')).to.be.true;
    }
  });
  test("viewFile should open a file if it's not an xml file", async function () {
    const commandsStub = stub(commands, 'executeCommand');
    const createDirectoryStub = stub(workspace.fs, 'createDirectory').callsFake(uri => {
      expect(uri).to.be.instanceof(Uri);
      return Promise.resolve();
    });
    const spawnStub = stub(child_process, 'spawn').returns({ stderr: null } as child_process.ChildProcess);
    const enc = new TextEncoder();
    const codeCat = enc.encode('tacocat');
    const writeFileStub = stub(workspace.fs, 'writeFile').callsFake((uri: Uri, u8a: Uint8Array) => {
      expect(uri).to.be.instanceof(Uri);
      expect(u8a).to.eq(codeCat);
      return Promise.resolve();
    });
    const zipStub = stub(ooxmlViewer.zip, 'file').returns(({
      async(arg: string) {
        if (arg === 'text') {
          return 'tacocat';
        } else if (arg === 'uint8array') {
          return Promise.resolve(codeCat);
        }
      },
    } as never) as JSZip);
    stubs.push(commandsStub, createDirectoryStub, spawnStub, writeFileStub, zipStub);
    const node = new FileNode();
    await ooxmlViewer.viewFile(node);
    const folderPath = join(OOXMLViewer.fileCachePath, dirname(node.fullPath));
    const filePath: string = join(folderPath, node.fileName);
    expect(OOXMLViewer.openTextEditors[filePath]).to.eq(node);
    expect(commandsStub.calledWith('vscode.open')).to.be.true;
    if (process && (process.platform === 'win32' || (process.env && process.env.OSTYPE && /^(msys|cygwin)$/.test(process.env.OSTYPE)))) {
      expect(spawnStub.calledWith('attrib')).to.be.true;
    }
  });
});
