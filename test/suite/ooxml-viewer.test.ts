import { expect } from 'chai';
import { tmpdir } from 'os';
import { join } from 'path';
import { createStubInstance, SinonStub, stub } from 'sinon';
import { ExtensionContext } from 'vscode';
import { OOXMLExtensionSettings } from '../../src/ooxml-extension-settings';
import { OOXMLPackageFacade } from '../../src/ooxml-package/ooxml-package-facade';
import { OOXMLViewer } from '../../src/ooxml-viewer';
import { FileNode, OOXMLTreeDataProvider } from '../../src/tree-view/ooxml-tree-view-provider';
import { ExtensionUtilities } from '../../src/utilities/extension-utilities';
import { FileSystemUtilities } from '../../src/utilities/file-system-utilities';

suite('OOXMLViewer', async function () {
  this.timeout(10000);
  let ooxmlViewer: OOXMLViewer;
  const stubs: SinonStub[] = [];
  const testFilePath = join(__dirname, '..', '..', '..', 'test', 'test-data', 'Test.pptx');
  const settings = <OOXMLExtensionSettings>{
    maximumOOXMLFileSizeBytes: 50000000,
  };

  setup(function () {
    const context = {
      storageUri: {
        fsPath: join(tmpdir(), 'ooxml-viewer'),
      },
      subscriptions: [],
    } as unknown as ExtensionContext;
    const treeViewDataProvider = createStubInstance(OOXMLTreeDataProvider);
    treeViewDataProvider.rootFileNode = new FileNode();
    ooxmlViewer = new OOXMLViewer(treeViewDataProvider, settings, context);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('openOOXMLPackage should create ooxml package', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(ooxmlPackage);
    stubs.push(createPackageStub);

    await ooxmlViewer.openOOXMLPackage(testFilePath);

    expect(createPackageStub.callCount).to.be.eq(1);
  });

  test('openOOXMLPackage should remove and recreate ooxml package if it is called twice on the same file', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(ooxmlPackage);
    stubs.push(createPackageStub);

    await ooxmlViewer.openOOXMLPackage(testFilePath);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    expect(createPackageStub.callCount).to.be.eq(2);
    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
  });

  test('openOOXMLPackage should create ooxml package', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(ooxmlPackage);
    const getFileSizeStub = stub(FileSystemUtilities, 'getFileSize').returns(Promise.resolve(100000000));
    const showWarningStub = stub(ExtensionUtilities, 'showWarning').returns(Promise.resolve());
    stubs.push(createPackageStub, getFileSizeStub, showWarningStub);

    await ooxmlViewer.openOOXMLPackage(testFilePath);

    expect(createPackageStub.callCount).to.be.eq(0);
    expect(showWarningStub.callCount).to.be.eq(1);
  });

  test('removeOOXMLPackage should dispose ooxml package', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(ooxmlPackage);
    stubs.push(createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    await ooxmlViewer.removeOOXMLPackage(testFilePath);

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
  });

  test('removeOOXMLPackage should handle being called twice', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    ooxmlPackage.ooxmlFilePath = testFilePath;
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(ooxmlPackage);
    stubs.push(createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);
    await ooxmlViewer.removeOOXMLPackage(testFilePath);

    await ooxmlViewer.removeOOXMLPackage(testFilePath);

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
  });

  test('reset should reset all packages and clear cache', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    const deleteFileStub = stub(FileSystemUtilities, 'deleteFile').returns(Promise.resolve());
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(ooxmlPackage);
    stubs.push(deleteFileStub, createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    await ooxmlViewer.reset();

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
    expect(deleteFileStub.callCount).to.be.eq(1);
  });

  test('reset should not error if deleteFile throws', async function () {
    const ooxmlPackage = createStubInstance(OOXMLPackageFacade);
    const deleteFileStub = stub(FileSystemUtilities, 'deleteFile').throws(new Error());
    const createPackageStub = stub(OOXMLPackageFacade, 'create').returns(ooxmlPackage);
    stubs.push(deleteFileStub, createPackageStub);
    await ooxmlViewer.openOOXMLPackage(testFilePath);

    await ooxmlViewer.reset();

    expect(ooxmlPackage.dispose.callCount).to.be.eq(1);
    expect(deleteFileStub.callCount).to.be.eq(1);
  });

  test('reset should not call deleteFile if there is no storageUri', async function () {
    const context = {
      storageUri: null,
      subscriptions: [],
    } as unknown as ExtensionContext;
    const treeViewDataProvider = createStubInstance(OOXMLTreeDataProvider);
    treeViewDataProvider.rootFileNode = new FileNode();
    const ooxmlViewer = new OOXMLViewer(treeViewDataProvider, settings, context);
    const deleteFileStub = stub(FileSystemUtilities, 'deleteFile').returns(Promise.resolve());
    stubs.push(deleteFileStub);

    await ooxmlViewer.reset();

    expect(deleteFileStub.callCount).to.be.eq(0);
  });
});
