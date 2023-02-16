import { expect } from 'chai';
import { tmpdir } from 'os';
import { join } from 'path';
import { createStubInstance, SinonStub, stub } from 'sinon';
import { ExtensionContext } from 'vscode';
import { OOXMLPackageFacade } from '../../src/ooxml-package/ooxml-package-facade';
import { OOXMLViewer } from '../../src/ooxml-viewer';
import { OOXMLTreeDataProvider } from '../../src/tree-view/ooxml-tree-view-provider';
import { FileSystemUtilities } from '../../src/utilities/file-system-utilities';

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
    const treeViewDataProvider = createStubInstance(OOXMLTreeDataProvider);
    ooxmlViewer = new OOXMLViewer(treeViewDataProvider, context);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('openOOXMLPackage should create ooxml package', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(Promise.resolve(ooxmlPackage));
    stubs.push(createPackageStub);

    await ooxmlViewer.openOOXMLPackage(testFilePath);

    expect(createPackageStub.callCount).to.be.eq(1);
  });

  test('openOOXMLPackage should remove and recreate ooxml package if it is called twice on the same file', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(Promise.resolve(ooxmlPackage));
    stubs.push(createPackageStub);

    await ooxmlViewer.openOOXMLPackage(testFilePath);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    expect(createPackageStub.callCount).to.be.eq(2);
    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
  });

  test('removeOOXMLPackage should dispose ooxml package', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(Promise.resolve(ooxmlPackage));
    stubs.push(createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    await ooxmlViewer.removeOOXMLPackage(testFilePath);

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
  });

  test('removeOOXMLPackage should handle being called twice', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(Promise.resolve(ooxmlPackage));
    stubs.push(createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);
    await ooxmlViewer.removeOOXMLPackage(testFilePath);

    await ooxmlViewer.removeOOXMLPackage(testFilePath);

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
  });

  test('reset should reset all packages and clear cache', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    const deleteFileStub = stub(FileSystemUtilities, 'deleteFile').returns(Promise.resolve());
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(Promise.resolve(ooxmlPackage));
    stubs.push(deleteFileStub, createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    await ooxmlViewer.reset();

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
    expect(deleteFileStub.callCount).to.be.eq(1);
  });

  test('reset should not error if deleteFile throws', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    const deleteFileStub = stub(FileSystemUtilities, 'deleteFile').throws(new Error());
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(Promise.resolve(ooxmlPackage));
    stubs.push(deleteFileStub, createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    await ooxmlViewer.reset();

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
    expect(deleteFileStub.callCount).to.be.eq(1);
  });
});
