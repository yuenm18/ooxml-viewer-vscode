import { expect } from 'chai';
import JSZip from 'jszip';
import { tmpdir } from 'os';
import { join, sep } from 'path';
import { SinonStub, spy, stub } from 'sinon';
import { TextDecoder } from 'util';
import { commands, Disposable, ExtensionContext, FileSystemError, Uri, window } from 'vscode';
import xmlFormatter from 'xml-formatter';
import { NORMAL_SUBFOLDER_NAME, OOXMLFileCache } from '../../src/ooxml-file-cache';
import { OOXMLPackage } from '../../src/ooxml-package';
import { FileNode, OOXMLTreeDataProvider } from '../../src/ooxml-tree-view-provider';
import { OOXMLViewer } from '../../src/ooxml-viewer';

suite('OOXMLPackage', async function () {
  this.timeout(10000);
  let ooxmlPackage: OOXMLPackage;
  const stubs: SinonStub[] = [];
  const testFilePath = join(__dirname, '..', '..', '..', 'test', 'test-data', 'Test.pptx');
  const treeView = new OOXMLTreeDataProvider();
  suiteSetup(function () {
    const context = {
      storageUri: {
        fsPath: join(tmpdir(), 'ooxml-viewer'),
      },
      subscriptions: [],
    } as unknown as ExtensionContext;

    ooxmlPackage = new OOXMLPackage(Uri.parse(testFilePath), treeView, context);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('searchOoxmlParts should return and not perform a search if no search term is entered', function (done) {
    const showInputStub = stub(window, 'showInputBox').returns(Promise.resolve(''));
    const formatXmlStub = stub(ooxmlPackage, <never>'formatXml');
    const executeCommandStub = stub(commands, 'executeCommand');
    stubs.push(showInputStub, formatXmlStub, executeCommandStub);

    ooxmlPackage.searchOoxmlParts().then(() => {
      expect(formatXmlStub.callCount).to.eq(0);
      expect(executeCommandStub.callCount).to.eq(0);
      done();
    });
  });

  test('should populate the sidebar tree with the contents of an ooxml file', async function () {
    const writeFileMock = stub(ooxmlPackage.cache, 'createCachedFiles').returns(Promise.resolve());
    const refreshStub = stub(this.treeDataProvider, 'refresh').returns(undefined);
    const jsZipStub = stub(ooxmlPackage.zip, 'file').callsFake(() => {
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
      stub(ooxmlPackage, <never>'hasFileBeenChangedFromOutside').returns(Promise.resolve(false)),
      jsZipStub,
      refreshStub,
      writeFileMock,
    );

    expect(this.treeDataProvider.rootFileNode.children[0].children.length).to.eq(0);
    await ooxmlPackage.openOoxmlPackage(Uri.file(testFilePath));

    expect(this.treeDataProvider.rootFileNode.children[0].children.length).to.eq(4);
    expect(refreshStub.callCount).to.eq(3);
    expect(writeFileMock.callCount).to.eq(40);
  });

  test('viewFile should open a text editor when called with the path to an xml file', async function () {
    const commandsStub = stub(commands, 'executeCommand');
    const enc = new TextEncoder();
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
      "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>";
    const coded = enc.encode(xml);
    const readCachedFileStub = stub(ooxmlPackage.cache, 'getCachedNormalFile').callsFake((path: string) => {
      return Promise.resolve(coded);
    });
    const readCompareFileStub = stub(ooxmlPackage.cache, 'getCachedCompareFile').callsFake((path: string) => {
      return Promise.resolve(coded);
    });
    const updateCachedFilesStub = stub(ooxmlPackage.cache, 'updateCachedFilesNoCompare').callsFake((path: string, u8a: Uint8Array) => {
      const sent: Buffer = Buffer.from(enc.encode(xmlFormatter(xml, { indentation: '  ', collapseContent: true })));
      const received: Buffer = Buffer.from(u8a);
      expect(sent.equals(received)).to.be.true;
      return Promise.resolve();
    });
    const updateCompareFileStub = stub(ooxmlPackage.cache, 'updateCompareFile').callsFake((path: string, u8a: Uint8Array) => {
      const sent: Buffer = Buffer.from(enc.encode(xmlFormatter(xml, { indentation: '  ', collapseContent: true })));
      const received: Buffer = Buffer.from(u8a);
      expect(sent.equals(received)).to.be.true;
      return Promise.resolve();
    });
    stubs.push(commandsStub, readCachedFileStub, readCompareFileStub, updateCachedFilesStub, updateCompareFileStub);
    const node = new FileNode();

    await ooxmlPackage.viewFile(node);

    expect(commandsStub.calledWith('vscode.open')).to.be.true;
  });

  test("viewFile should open a file if it's not an xml file", async function () {
    const commandsStub = stub(commands, 'executeCommand');
    const enc = new TextEncoder();
    const codeCat = enc.encode('tacocat');
    const updateCachedFilesStub = stub(ooxmlPackage.cache, 'updateCachedFilesNoCompare').callsFake((path: string, u8a: Uint8Array) => {
      expect(u8a).to.eq(codeCat);
      return Promise.resolve();
    });
    const updateCompareFileStub = stub(ooxmlPackage.cache, 'updateCompareFile').callsFake((path: string, u8a: Uint8Array) => {
      expect(u8a).to.eq(codeCat);
      return Promise.resolve();
    });
    const readCachedFileStub = stub(ooxmlPackage.cache, 'getCachedNormalFile').returns(Promise.resolve(new Uint8Array()));
    const readCompareFileStub = stub(ooxmlPackage.cache, 'getCachedCompareFile').returns(Promise.resolve(new Uint8Array()));
    stubs.push(commandsStub, updateCachedFilesStub, updateCompareFileStub, readCachedFileStub, readCompareFileStub);
    const node = new FileNode();

    await ooxmlPackage.viewFile(node);

    expect(commandsStub.calledWith('vscode.open')).to.be.true;
  });

  test('clear should reset the OOXML Viewer', async function () {
    const refreshStub = stub(this.treeDataProvider, 'refresh').callsFake(() => undefined);
    const disposeWatchersStub = stub(OOXMLViewer.prototype, 'disposeWatchers').callsFake(() => undefined);
    const clearCacheStub = stub(OOXMLFileCache.prototype, 'reset').callsFake(() => Promise.resolve());

    stubs.push(refreshStub, disposeWatchersStub, clearCacheStub);

    await ooxmlPackage.clear();

    expect(refreshStub.calledOnce).to.be.true;
    expect(disposeWatchersStub.calledOnce).to.be.true;
    expect(clearCacheStub.calledOnce).to.be.true;
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
    const readCachedFileStub = stub(ooxmlPackage.cache, 'getCachedNormalFile').returns(Promise.resolve(new TextEncoder().encode(xml)));
    const readCompareFileStub = stub(ooxmlPackage.cache, 'getCachedCompareFile').returns(Promise.resolve(new TextEncoder().encode(xml)));
    const updateCachedFilesStub = stub(ooxmlPackage.cache, 'updateCachedFilesNoCompare').callsFake((arg1, arg2) => {
      const dec = new TextDecoder();
      expect(dec.decode(arg2)).to.eq(xmlFormatter(xml, { indentation: '  ', collapseContent: true }));
      return Promise.resolve();
    });
    const updateCompareFileStub = stub(ooxmlPackage.cache, 'updateCompareFile').callsFake((arg1, arg2) => {
      const dec = new TextDecoder();
      expect(dec.decode(arg2)).to.eq(xmlFormatter(xml, { indentation: '  ', collapseContent: true }));
      return Promise.resolve();
    });
    const textDecoderStub = stub(ooxmlPackage.textDecoder, 'decode').callsFake(arg => {
      return xml;
    });
    stubs.push(vscodeDiffStub, readCachedFileStub, readCompareFileStub, updateCachedFilesStub, updateCompareFileStub, textDecoderStub);
    const node = new FileNode();
    node.fullPath = 'tacocat/racecar.xml';
    node.fileName = 'racecar.xml';

    await ooxmlPackage.getDiff(node);
  });

  test('getDiff should call display an error message when an error is thrown', async function () {
    const err = new Error('Pants on backwards');
    const encoderStub = stub(ooxmlPackage.textDecoder, 'decode').throws(err);
    const getCachedNormalFileStub = stub(ooxmlPackage.cache, 'getCachedNormalFile').returns(Promise.resolve(new Uint8Array()));
    const showErrorStub = stub(window, 'showErrorMessage');

    stubs.push(encoderStub, showErrorStub, getCachedNormalFileStub);

    await ooxmlPackage.getDiff('path');

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
    ooxmlPackage.watchers.push(disposable1, disposable2);

    ooxmlPackage.disposeWatchers();

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
    await ooxmlPackage.openOoxmlPackage(Uri.file(testFilePath));
    const populateOOXMLViewerStub = stub(ooxmlPackage, <never>'populateOOXMLViewer').callThrough();
    const getNormalFileCachePathStub = stub(ooxmlPackage.cache, 'getNormalFileCachePath');
    const updateCachedFilesStub = stub(ooxmlPackage.cache, 'updateCachedFiles');
    const deleteCachedFilesFileStub = stub(ooxmlPackage.cache, 'deleteCachedFiles');
    const readFileStub = stub(ooxmlPackage.cache, <never>'readFile').returns(Promise.resolve(new Uint8Array(8)));
    stubs.push(populateOOXMLViewerStub, deleteCachedFilesFileStub, updateCachedFilesStub, getNormalFileCachePathStub, readFileStub);
    delete ooxmlPackage.zip.files[path];
    const node = findNode(ooxmlPackage.treeDataProvider.rootFileNode, path);

    node?.setDeleted();

    await populateOOXMLViewerStub.bind(ooxmlPackage)(ooxmlPackage.zip.files, false);

    expect(deleteCachedFilesFileStub.called).to.be.true;
  });

  test('It should set a file node as deleted if it is removed from this.zip.files', async function () {
    const path = 'ppt/slides/slide1.xml';
    await ooxmlPackage.openOoxmlPackage(Uri.file(testFilePath));
    const populateOOXMLViewerStub = stub(ooxmlPackage, <never>'populateOOXMLViewer').callThrough();
    const updateCachedFilesStub = stub(ooxmlPackage.cache, 'updateCachedFiles');
    const deleteCachedFilesFileStub = stub(ooxmlPackage.cache, 'deleteCachedFiles');
    stubs.push(populateOOXMLViewerStub, deleteCachedFilesFileStub, updateCachedFilesStub);

    delete ooxmlPackage.zip.files[path];

    const node = findNode(this.treeDataProvider.rootFileNode.children[0], path);
    await populateOOXMLViewerStub.bind(ooxmlPackage)(ooxmlPackage.zip.files, false);
    expect(deleteCachedFilesFileStub.called).to.be.false;
    expect(node?.isDeleted()).to.be.true;
  });

  test('searchOoxmlParts should show an input box and use the input to perform a search of the OOXML parts', async function () {
    const searchTerm = 'meatballs';
    const showInputStub = stub(window, 'showInputBox').returns(Promise.resolve(searchTerm));
    const executeCommandStub = stub(commands, 'executeCommand');

    stubs.push(showInputStub, executeCommandStub);

    await ooxmlPackage.searchOoxmlParts();

    expect(executeCommandStub.args[0][0]).to.eq('workbench.action.findInFiles');
    expect(executeCommandStub.args[0][1]).to.deep.eq({
      query: searchTerm,
      filesToInclude: ooxmlPackage.cache.normalSubfolderPath,
      triggerSearch: true,
      isCaseSensitive: false,
      matchWholeWord: false,
    });
  });

  test('searchOoxmlParts should call window.showErrorMessage with an error if an error is thrown', function (done) {
    const err = new Error('out of tacos');
    const showInputStub = stub(window, 'showInputBox').throws(err);
    const showErrorMessageStub = stub(window, 'showErrorMessage');

    stubs.push(showInputStub, showErrorMessageStub);

    ooxmlPackage
      .searchOoxmlParts()
      .then(() => {
        expect(showErrorMessageStub.args[0][0]).to.eq(err.message);
        done();
      })
      .catch(error => {
        done(error);
      });
  });

  test('searchOoxmlParts should window.showWarningMessage if no file is open in the viewer with FileNotFound', function (done) {
    const msg = 'A file must be open in the OOXML Viewer to search its parts.';
    const showInputStub = stub(window, 'showInputBox').throws(FileSystemError.FileNotFound);
    const showWarningMessageStub = stub(window, 'showWarningMessage');
    stubs.push(showInputStub, showWarningMessageStub);

    ooxmlPackage.searchOoxmlParts().then(() => {
      expect(showWarningMessageStub.args[0][0]).to.eq(msg);
      done();
    });
  });

  test('tryFormatDocument should format document if it belongs to the normal cache path', async () => {
    const filePath = `${ooxmlPackage.cache.cacheBasePath}${sep}${NORMAL_SUBFOLDER_NAME}${sep}test`;
    const formatXmlStub = stub(ooxmlPackage, <never>'formatXml');
    stubs.push(formatXmlStub);

    await ooxmlPackage.tryFormatDocument(filePath);

    expect(formatXmlStub.called).to.be.true;
  });

  test('tryFormatDocument should not format document if it does not belongs to the normal cache path', async () => {
    const filePath = `random${sep}test`;
    const formatXmlStub = stub(ooxmlPackage, <never>'formatXml');
    stubs.push(formatXmlStub);

    await ooxmlPackage.tryFormatDocument(filePath);

    expect(formatXmlStub.called).to.be.false;
  });
});
