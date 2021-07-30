import { expect } from 'chai';
import findInFiles, { FindResult } from 'find-in-files';
import JSZip from 'jszip';
import { tmpdir } from 'os';
import { join, sep } from 'path';
import { match, SinonStub, spy, stub } from 'sinon';
import { TextDecoder } from 'util';
import { commands, Disposable, ExtensionContext, TextDocument, TextDocumentShowOptions, TextEditor, Uri, window } from 'vscode';
import xmlFormatter from 'xml-formatter';
import { CACHE_FOLDER_NAME, NORMAL_SUBFOLDER_NAME } from '../../src/ooxml-file-cache';
import { FileNode, OOXMLTreeDataProvider } from '../../src/ooxml-tree-view-provider';
import { OOXMLViewer } from '../../src/ooxml-viewer';

suite('OOXMLViewer', async function () {
  this.timeout(10000);
  let ooxmlViewer: OOXMLViewer;
  const stubs: SinonStub[] = [];
  const testFilePath = join(__dirname, '..', '..', '..', 'test', 'test-data', 'Test.pptx');

  setup(function () {
    const context = {
      storageUri: {
        fsPath: join(tmpdir(), 'ooxml-viewer'),
      },
      subscriptions: [],
    } as unknown as ExtensionContext;
    ooxmlViewer = new OOXMLViewer(context);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('should have an instance of OOXMLTreeDataProvider', function () {
    expect(ooxmlViewer.treeDataProvider).to.be.instanceOf(OOXMLTreeDataProvider);
  });

  test('should populate the sidebar tree with the contents of an ooxml file', async function () {
    const writeFileMock = stub(ooxmlViewer.cache, 'writeFile').returns(Promise.resolve());
    const refreshStub = stub(ooxmlViewer.treeDataProvider, 'refresh').returns(undefined);
    const jsZipStub = stub(ooxmlViewer.zip, 'file').callsFake(() => {
      return {
        async(arg: string) {
          expect(arg).to.eq('text');
          return Promise.resolve(
            '<?xml version="1.0" encoding="UTF-8"?>\n<note><date>2015-09-01</date><hour>08:30</hour>' +
              "<to>Tove</to><from>Jani</from><body>Don't forget me this weekend!</body></note>",
          );
        },
      } as unknown as JSZip;
    });
    stubs.push(
      stub(ooxmlViewer, <never>'hasFileBeenChangedFromOutside').returns(Promise.resolve(false)),
      jsZipStub,
      refreshStub,
      writeFileMock,
    );

    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(0);
    await ooxmlViewer.openOoxmlPackage(Uri.file(testFilePath));

    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(4);
    expect(refreshStub.callCount).to.eq(3);
    expect(writeFileMock.callCount).to.eq(120);
  });

  test('viewFile should open a text editor when called with the path to an xml file', async function () {
    const commandsStub = stub(commands, 'executeCommand');
    const enc = new TextEncoder();
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
      "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>";
    const coded = enc.encode(xml);
    const readFileStub = stub(ooxmlViewer.cache, 'readFile').callsFake((path: string) => {
      return Promise.resolve(coded);
    });
    const writeFileStub = stub(ooxmlViewer.cache, 'writeFile').callsFake((path: string, u8a: Uint8Array) => {
      const sent: Buffer = Buffer.from(enc.encode(xmlFormatter(xml, { indentation: '  ', collapseContent: true })));
      const received: Buffer = Buffer.from(u8a);
      expect(sent.equals(received)).to.be.true;
      return Promise.resolve();
    });
    stubs.push(commandsStub, readFileStub, writeFileStub);
    const node = new FileNode();

    await ooxmlViewer.viewFile(node);

    const filePath = ooxmlViewer.cache.getFileCachePath(node.fullPath);
    expect(ooxmlViewer.openTextEditors[filePath]).to.eq(node);
    expect(commandsStub.calledWith('vscode.open')).to.be.true;
  });

  test("viewFile should open a file if it's not an xml file", async function () {
    const commandsStub = stub(commands, 'executeCommand');
    const enc = new TextEncoder();
    const codeCat = enc.encode('tacocat');
    const writeFileStub = stub(ooxmlViewer.cache, 'writeFile').callsFake((path: string, u8a: Uint8Array) => {
      expect(u8a).to.eq(codeCat);
      return Promise.resolve();
    });
    const readFileStub = stub(ooxmlViewer.cache, 'readFile').returns(Promise.resolve(new Uint8Array()));
    stubs.push(commandsStub, writeFileStub, readFileStub);
    const node = new FileNode();

    await ooxmlViewer.viewFile(node);

    expect(commandsStub.calledWith('vscode.open')).to.be.true;
  });

  test('clear should reset the OOXML Viewer', async function () {
    const textDoc = {} as TextDocument;
    const refreshStub = stub(ooxmlViewer.treeDataProvider, 'refresh').callsFake(() => undefined);
    const disposeWatchersStub = stub(OOXMLViewer.prototype, 'disposeWatchers').callsFake(() => undefined);
    const openEditorsStub = stub(Array.prototype, 'filter').callsFake(arg => {
      return [textDoc];
    });
    const showDocStub = stub(window, 'showTextDocument').callsFake((td: Uri, config?: TextDocumentShowOptions | undefined) => {
      expect(td).to.eq(textDoc);
      expect(config).to.deep.eq({ preview: true, preserveFocus: false });
      return Promise.resolve({} as TextEditor);
    });
    const executeStub = stub(commands, 'executeCommand').callsFake(arg => {
      return Promise.resolve();
    });
    stubs.push(refreshStub, disposeWatchersStub, openEditorsStub, showDocStub, executeStub);

    await ooxmlViewer.clear();

    expect(refreshStub.calledOnce).to.be.true;
    expect(disposeWatchersStub.calledOnce).to.be.true;
    expect(showDocStub.calledWith(match(textDoc))).to.be.true;
    expect(executeStub.args[0][0]).eq('workbench.action.closeActiveEditor');
  });

  test('getDiff should use vscode.diff to get the difference between two files', async function () {
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
      "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>";
    const vscodeDiffStub = stub(commands, 'executeCommand').callsFake((cmd: string, leftUri, rightUri, title) => {
      expect(cmd).to.eq('vscode.diff');
      expect(title).to.eq('racecar.xml ↔ compare.racecar.xml');
      expect(leftUri).to.be.instanceof(Uri);
      expect(rightUri).to.be.instanceof(Uri);
      expect(leftUri.path).to.include('compare');
      expect(rightUri.path).to.include('racecar.xml');
      expect(rightUri.path).not.to.include('compare');
      return Promise.resolve();
    });
    const readFileStub = stub(ooxmlViewer.cache, 'readFile').returns(
      Promise.resolve({
        toString() {
          return xml;
        },
      } as Uint8Array),
    );
    const writeFileStub = stub(ooxmlViewer.cache, 'writeFile').callsFake((arg1, arg2) => {
      expect(arg1).to.include(CACHE_FOLDER_NAME);
      const dec = new TextDecoder();
      expect(dec.decode(arg2)).to.eq(xmlFormatter(xml, { indentation: '  ', collapseContent: true }));
      return Promise.resolve();
    });
    const textDecoderStub = stub(ooxmlViewer.textDecoder, 'decode').callsFake(arg => {
      return xml;
    });
    stubs.push(vscodeDiffStub, readFileStub, writeFileStub, textDecoderStub);
    const node = new FileNode();
    node.fullPath = 'tacocat/racecar.xml';
    node.fileName = 'racecar.xml';

    await ooxmlViewer.getDiff(node);
  });

  test('getDiff should call display an error message when an error is thrown', async function () {
    const err = new Error('Pants on backwards');
    const encoderStub = stub(ooxmlViewer.textDecoder, 'decode').throws(err);
    const getCachedFileStub = stub(ooxmlViewer.cache, 'getCachedFile').returns(Promise.resolve(new Uint8Array()));
    const consoleErrorStub = stub(console, 'error');
    const showErrorStub = stub(window, 'showErrorMessage');

    stubs.push(encoderStub, consoleErrorStub, showErrorStub, getCachedFileStub);

    await ooxmlViewer.getDiff(new FileNode());

    expect(consoleErrorStub.args[0][0]).to.eq(err.message);
    expect(showErrorStub.args[0][0]).to.eq(err.message);
  });

  test('closeWatchers should call restore on the array of file system watchers', function () {
    const disposeStub = spy();
    const disposable1 = {
      dispose: disposeStub,
    } as never as Disposable;
    const disposable2 = {
      dispose: disposeStub,
    } as never as Disposable;
    ooxmlViewer.watchers.push(disposable1, disposable2);

    ooxmlViewer.disposeWatchers();

    expect(disposeStub.calledTwice).to.be.true;
  });

  function findNode(node: FileNode, path: string): FileNode | null {
    let found: FileNode | null = null;
    if (node.fullPath === path) {
      return node;
    } else if (node.children.length) {
      for (let i = 0; !found && i < node.children.length; i++) {
        found = findNode(node.children[i], path);
      }
    }
    return found;
  }

  test('It should delete a file node if isDeleted returns true', async function () {
    const path = 'ppt/slides/slide1.xml';
    await ooxmlViewer.openOoxmlPackage(Uri.file(testFilePath));
    const populateOOXMLViewerStub = stub(ooxmlViewer, <never>'populateOOXMLViewer').callThrough();
    const getFileCachePathStub = stub(ooxmlViewer.cache, 'getFileCachePath');
    const updateCachedFileStub = stub(ooxmlViewer.cache, 'updateCachedFile');
    const deleteCachedFilesFileStub = stub(ooxmlViewer.cache, 'deleteCachedFiles');
    const readFileStub = stub(ooxmlViewer.cache, <never>'readFile').returns(Promise.resolve(new Uint8Array(8)));
    stubs.push(populateOOXMLViewerStub, deleteCachedFilesFileStub, updateCachedFileStub, getFileCachePathStub, readFileStub);
    delete ooxmlViewer.zip.files[path];
    const node = findNode(ooxmlViewer.treeDataProvider.rootFileNode, path);

    node?.setDeleted();

    await populateOOXMLViewerStub.bind(ooxmlViewer)(ooxmlViewer.zip.files, false);

    expect(deleteCachedFilesFileStub.called).to.be.true;
  });

  test('It should set a file node as deleted if it is removed from this.zip.files', async function () {
    const path = 'ppt/slides/slide1.xml';
    await ooxmlViewer.openOoxmlPackage(Uri.file(testFilePath));
    const populateOOXMLViewerStub = stub(ooxmlViewer, <never>'populateOOXMLViewer').callThrough();
    const updateCachedFileStub = stub(ooxmlViewer.cache, 'updateCachedFile');
    const deleteCachedFilesFileStub = stub(ooxmlViewer.cache, 'deleteCachedFiles');
    stubs.push(populateOOXMLViewerStub, deleteCachedFilesFileStub, updateCachedFileStub);

    delete ooxmlViewer.zip.files[path];

    const node = findNode(ooxmlViewer.treeDataProvider.rootFileNode, path);
    await populateOOXMLViewerStub.bind(ooxmlViewer)(ooxmlViewer.zip.files, false);
    expect(deleteCachedFilesFileStub.called).to.be.false;
    expect(node?.isDeleted()).to.be.true;
  });

  test('searchOoxmlParts should return and not perform a search if no search term is entered', async function () {
    const showInputStub = stub(window, 'showInputBox').returns(Promise.resolve(''));
    const tryFormatXmlStub = stub(ooxmlViewer, <never>'tryFormatXml');
    const executeCommandStub = stub(commands, 'executeCommand');
    const findStub = stub(findInFiles, 'find');
    stubs.push(showInputStub, tryFormatXmlStub, executeCommandStub, findStub);

    await ooxmlViewer.searchOxmlParts();
    expect(tryFormatXmlStub.callCount).to.eq(0);
    expect(executeCommandStub.callCount).to.eq(0);
    expect(findStub.callCount).to.eq(0);
  });

  test('searchOoxmlParts should show an input box and use the input to perform a search of the OOXML parts', async function () {
    const searchTerm = 'meatballs';
    const filePath0 = `helloworld${sep + NORMAL_SUBFOLDER_NAME + sep}racecar${sep}radar`;
    const filePath1 = `foo${sep + NORMAL_SUBFOLDER_NAME + sep}bar${sep}baz`;
    const filePath2 = `I${sep + NORMAL_SUBFOLDER_NAME + sep}want${sep}breakfast`;
    const showInputStub = stub(window, 'showInputBox').returns(Promise.resolve(searchTerm));
    const tryFormatXmlStub = stub(ooxmlViewer, <never>'tryFormatXml');
    const executeCommandStub = stub(commands, 'executeCommand');
    const findStub = stub(findInFiles, 'find').returns(
      Promise.resolve({
        [filePath0]: {},
        [filePath1]: {},
        [filePath2]: {},
      } as unknown as FindResult),
    );

    stubs.push(showInputStub, tryFormatXmlStub, executeCommandStub, findStub);

    await ooxmlViewer.searchOxmlParts();

    expect(findStub.args[0][0]).to.eq(searchTerm);
    expect(findStub.args[0][1]).to.eq(ooxmlViewer.cache.normalSubfolderPath);
    expect(tryFormatXmlStub.callCount).to.eq(3);
    expect(tryFormatXmlStub.args[0][0]).to.eq(filePath0.split(NORMAL_SUBFOLDER_NAME)[1].split(sep).join('/'));
    expect(tryFormatXmlStub.args[1][0]).to.eq(filePath1.split(NORMAL_SUBFOLDER_NAME)[1].split(sep).join('/'));
    expect(tryFormatXmlStub.args[2][0]).to.eq(filePath2.split(NORMAL_SUBFOLDER_NAME)[1].split(sep).join('/'));
    expect(executeCommandStub.args[0][0]).to.eq('workbench.action.findInFiles');
    expect(executeCommandStub.args[0][1]).to.deep.eq({
      query: searchTerm,
      filesToInclude: ooxmlViewer.cache.normalSubfolderPath,
      triggerSearch: true,
      isCaseSensitive: false,
      matchWholeWord: false,
    });
  });

  test('searchOoxmlParts should console.error and window.showErrorMessage an error if an error is thrown', async function () {
    const err = new Error('out of tacos');
    const showInputStub = stub(window, 'showInputBox').returns(Promise.reject(err));
    const consoleErrorStub = stub(console, 'error');
    const showErrorMessageStub = stub(window, 'showErrorMessage');
    stubs.push(showInputStub, consoleErrorStub, showErrorMessageStub);

    await ooxmlViewer.searchOxmlParts();
    expect(consoleErrorStub.args[0][0]).to.eq(err.message);
    expect(showErrorMessageStub.args[0][0]).to.eq(err.message);
  });

  class VSError extends Error {
    code: string;

    constructor(data: { code: string }) {
      super();
      this.code = data.code;
    }
  }

  test('searchOoxmlParts should window.showWarningMessage if no file is open in the viewer with ENOENT', async function () {
    const err = new VSError({ code: 'ENOENT' });
    const msg = 'A file must be open in the OOXML Viewer to search its parts.';
    const showInputStub = stub(window, 'showInputBox').returns(Promise.reject(err));
    const showWarningMessageStub = stub(window, 'showWarningMessage');
    stubs.push(showInputStub, showWarningMessageStub);

    await ooxmlViewer.searchOxmlParts();
    expect(showWarningMessageStub.args[0][0]).to.eq(msg);
  });

  test('searchOoxmlParts should window.showWarningMessage if no file is open in the viewer with FileNotFound', async function () {
    const err = new VSError({ code: 'FileNotFound' });
    const msg = 'A file must be open in the OOXML Viewer to search its parts.';
    const showInputStub = stub(window, 'showInputBox').returns(Promise.reject(err));
    const showWarningMessageStub = stub(window, 'showWarningMessage');
    stubs.push(showInputStub, showWarningMessageStub);

    await ooxmlViewer.searchOxmlParts();
    expect(showWarningMessageStub.args[0][0]).to.eq(msg);
  });
});
