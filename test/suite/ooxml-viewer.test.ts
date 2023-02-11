import { expect } from 'chai';
import JSZip from 'jszip';
import { tmpdir } from 'os';
import { join } from 'path';
import { SinonStub, stub } from 'sinon';
import { ExtensionContext, Uri } from 'vscode';
import { OOXMLFileCache } from '../../src/ooxml-file-cache';
import { OOXMLViewer } from '../../src/ooxml-viewer';

suite('OOXMLViewer', async function () {
  this.timeout(10000);
  let ooxmlViewer: OOXMLViewer;
  const stubs: SinonStub[] = [];
  const testFilePath = join(__dirname, '..', '..', '..', 'test', 'test-data', 'Test.pptx');

  suiteSetup(function () {
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

  test('reset should reset the OOXML Viewer', async function () {
    const refreshStub = stub(ooxmlViewer.treeDataProvider, 'refresh').callsFake(() => undefined);
    const disposeWatchersStub = stub(OOXMLViewer.prototype, 'reset').callsFake(() => Promise.resolve());
    const clearCacheStub = stub(OOXMLFileCache.prototype, 'reset').callsFake(() => Promise.resolve());

    stubs.push(refreshStub, disposeWatchersStub, clearCacheStub);

    await ooxmlViewer.reset();

    expect(refreshStub.calledOnce).to.be.true;
    expect(disposeWatchersStub.calledOnce).to.be.true;
    expect(clearCacheStub.calledOnce).to.be.true;
  });

  test('openOoxmlPackage opens package ', async function () {
    const writeFileMock = stub(ooxmlPackage.cache, 'createCachedFiles').returns(Promise.resolve());
    const refreshStub = stub(this.treeDataProvider, 'refresh').returns(undefined);
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

    expect(this.treeDataProvider.rootFileNode.children[0].children.length).to.eq(0);
    await ooxmlViewer.openOoxmlPackage(Uri.file(testFilePath));

    expect(this.treeDataProvider.rootFileNode.children[0].children.length).to.eq(4);
    expect(refreshStub.callCount).to.eq(3);
    expect(writeFileMock.callCount).to.eq(40);
  });
});
