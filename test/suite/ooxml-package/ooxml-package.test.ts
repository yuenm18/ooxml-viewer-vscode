import { expect } from 'chai';
import { join } from 'path';
import { createStubInstance, SinonStub, SinonStubbedInstance, stub } from 'sinon';
import { ThemeIcon, TreeItemCollapsibleState, Uri } from 'vscode';
import { OOXMLExtensionSettings } from '../../../src/ooxml-extension-settings';
import { OOXMLPackage } from '../../../src/ooxml-package/ooxml-package';
import { OOXMLPackageFileAccessor } from '../../../src/ooxml-package/ooxml-package-file-accessor';
import { OOXMLPackageFileCache } from '../../../src/ooxml-package/ooxml-package-file-cache';
import { OOXMLPackageTreeView } from '../../../src/ooxml-package/ooxml-package-tree-view';
import { FileNode } from '../../../src/tree-view/ooxml-tree-view-provider';
import { ExtensionUtilities } from '../../../src/utilities/extension-utilities';
import { FileSystemUtilities } from '../../../src/utilities/file-system-utilities';
import { XmlFormatter } from '../../../src/utilities/xml-formatter';

suite('OOXMLPackage', async function () {
  this.timeout(10000);

  let ooxmlPackage: OOXMLPackage;

  let ooxmlFilePath = join(__dirname, '..', '..', '..', 'test', 'test-data', 'Test.pptx');
  let ooxmlFileAccessor: SinonStubbedInstance<OOXMLPackageFileAccessor>;
  let ooxmlPackageTreeView: SinonStubbedInstance<OOXMLPackageTreeView>;
  let cache: SinonStubbedInstance<OOXMLPackageFileCache>;
  let extensionSettings: OOXMLExtensionSettings;

  const stubs: SinonStub[] = [];

  setup(function () {
    ooxmlFilePath = 'package.json';
    ooxmlFileAccessor = createStubInstance(OOXMLPackageFileAccessor);
    ooxmlPackageTreeView = createStubInstance(OOXMLPackageTreeView);
    cache = createStubInstance(OOXMLPackageFileCache);
    extensionSettings = <OOXMLExtensionSettings>{};

    ooxmlPackage = new OOXMLPackage(ooxmlFilePath, ooxmlFileAccessor, ooxmlPackageTreeView, cache, extensionSettings);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  suite('viewFile', async function () {
    test('should open a file using the cached file path', async function () {
      const commandsStub = stub(ExtensionUtilities, 'openFile');
      stubs.push(commandsStub);
      cache.getNormalFileCachePath.withArgs('file-path').returns('cached-file-path');

      await ooxmlPackage.viewFile('file-path');

      expect(commandsStub.callCount).to.be.equal(1);
      expect(commandsStub.args[0][0]).to.be.equal('cached-file-path');
    });

    test('should format the cached file on open', async function () {
      const commandsStub = stub(ExtensionUtilities, 'openFile');
      stubs.push(commandsStub);
      const decoder = new TextDecoder();
      const encoder = new TextEncoder();
      cache.getCachedNormalFile.withArgs('file-path').returns(Promise.resolve(encoder.encode('<?xml ?><!-- comment --><Types></Types>')));

      await ooxmlPackage.viewFile('file-path');

      expect(cache.updateCachedFilesNoCompare.callCount).to.be.equal(1);
      expect(cache.updateCachedFilesNoCompare.args[0][0]).to.be.equal('file-path');
      expect(decoder.decode(cache.updateCachedFilesNoCompare.args[0][1])).to.be.equals('<?xml?>\r\n<!-- comment -->\r\n<Types></Types>');
    });

    test('should display error if an error is thrown', async function () {
      const showErrorStub = stub(ExtensionUtilities, 'showError');

      stubs.push(showErrorStub);
      cache.getNormalFileCachePath.throws(new Error());

      await ooxmlPackage.viewFile('file-path');

      expect(showErrorStub.callCount).to.be.equal(1);
    });
  });

  suite('getDiff', async function () {
    test('should use vscode.diff to get the difference between two files', async function () {
      const encoder = new TextEncoder();
      const decoder = new TextDecoder();
      const xml =
        '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
        "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>";
      const vscodeDiffStub = stub(ExtensionUtilities, 'openDiff').callsFake((leftPath, rightPath, title) => {
        expect(title).to.eq('racecar.xml ↔ compare.racecar.xml');
        expect(leftPath).to.eq('compare/racecar.xml');
        expect(rightPath).to.eq('normal/racecar.xml');
        return Promise.resolve();
      });
      cache.getCachedNormalFile.returns(Promise.resolve(encoder.encode(xml)));
      cache.getCachedCompareFile.returns(Promise.resolve(encoder.encode(xml)));
      cache.getNormalFileCachePath.returns('normal/racecar.xml');
      cache.getCompareFileCachePath.returns('compare/racecar.xml');

      stubs.push(vscodeDiffStub);

      const fullPath = 'tacocat/racecar.xml';

      await ooxmlPackage.getDiff(fullPath);

      expect(cache.updateCachedFilesNoCompare.callCount).to.equal(1);
      expect(cache.updateCachedFilesNoCompare.args[0][1]).to.deep.equal(XmlFormatter.format(encoder.encode(xml)));
      expect(cache.updateCompareFile.callCount).to.equal(1);
      expect(cache.updateCompareFile.args[0][1]).to.deep.equal(XmlFormatter.format(encoder.encode(xml)));
    });

    test('should call display an error message when an error is thrown', async function () {
      const err = new Error('Pants on backwards');
      cache.getCachedNormalFile.throws(err);
      const showErrorStub = stub(ExtensionUtilities, 'showError');

      stubs.push(showErrorStub);

      await ooxmlPackage.getDiff('path');

      expect((showErrorStub.args[0][0] as Error).message).to.eq(err.message);
    });
  });

  suite('tryFormatDocument', async function () {
    test('should format document if it belongs to the normal cache path', async () => {
      cache.cachePathIsNormal.returns(true);
      cache.getCachedNormalFile.returns(Promise.resolve(new TextEncoder().encode('<?xml ?><html></html>')));
      const filePath = 'file/path/document.xml';

      await ooxmlPackage.tryFormatDocument(filePath);

      expect(cache.updateCachedFilesNoCompare.callCount).to.be.equal(1);
    });

    test('should not format document if it does not belongs to the normal cache path', async () => {
      cache.cachePathIsNormal.returns(false);
      cache.getCachedNormalFile.returns(Promise.resolve(new TextEncoder().encode('<?xml ?><html></html>')));
      const filePath = 'file/path/document.xml';

      await ooxmlPackage.tryFormatDocument(filePath);

      expect(cache.updateCachedFilesNoCompare.callCount).to.be.equal(0);
    });

    test('should show error if error is thrown', async () => {
      cache.cachePathIsNormal.returns(true);
      const err = new Error('format error');
      cache.getCachedNormalFile.throws(err);
      const showErrorStub = stub(ExtensionUtilities, 'showError');
      stubs.push(showErrorStub);

      const filePath = 'file/path/document.xml';

      await ooxmlPackage.tryFormatDocument(filePath);

      expect(showErrorStub.callCount).to.eq(1);
      expect((showErrorStub.args[0][0] as Error).message).to.eq(err.message);
      expect(cache.updateCachedFilesNoCompare.callCount).to.be.equal(0);
    });
  });

  suite('searchOOXMLParts', async function () {
    test('should return and not perform a search if no search term is entered', async function () {
      const normalSubfolderPathStub = stub(cache, 'normalSubfolderPath').get(() => 'normal/path');
      const fileExistsStub = stub(FileSystemUtilities, 'fileExists').returns(Promise.resolve(true));
      const showInputStub = stub(ExtensionUtilities, 'showInput').returns(Promise.resolve(''));
      const findInFilesStub = stub(ExtensionUtilities, 'findInFiles').returns(Promise.resolve());
      stubs.push(normalSubfolderPathStub, showInputStub, fileExistsStub, findInFilesStub);

      await ooxmlPackage.searchOOXMLParts();

      expect(findInFilesStub.callCount).to.eq(0);
    });

    test('should show an input box and use the input to perform a search of the OOXML parts', async function () {
      const searchTerm = 'meatballs';
      const normalSubfolderPathStub = stub(cache, 'normalSubfolderPath').get(() => 'normal/path');
      const fileExistsStub = stub(FileSystemUtilities, 'fileExists').returns(Promise.resolve(true));
      const showInputStub = stub(ExtensionUtilities, 'showInput').returns(Promise.resolve(searchTerm));
      const findInFilesStub = stub(ExtensionUtilities, 'findInFiles').returns(Promise.resolve());
      stubs.push(normalSubfolderPathStub, showInputStub, fileExistsStub, findInFilesStub);

      await ooxmlPackage.searchOOXMLParts();

      expect(findInFilesStub.callCount).to.eq(1);
      expect(findInFilesStub.args[0][0]).to.deep.eq(searchTerm, 'normal/path');
    });

    test('should show error message if an error is thrown', function (done) {
      const err = new Error('out of tacos');
      const normalSubfolderPathStub = stub(cache, 'normalSubfolderPath').get(() => 'normal/path');
      const fileExistsStub = stub(FileSystemUtilities, 'fileExists').returns(Promise.resolve(true));
      const showInputStub = stub(ExtensionUtilities, 'showInput').throws(err);
      const findInFilesStub = stub(ExtensionUtilities, 'findInFiles').returns(Promise.resolve());
      const showErrorMessageStub = stub(ExtensionUtilities, 'showError').returns(Promise.resolve());
      stubs.push(normalSubfolderPathStub, showInputStub, fileExistsStub, findInFilesStub, showErrorMessageStub);

      ooxmlPackage
        .searchOOXMLParts()
        .then(() => {
          expect((showErrorMessageStub.args[0][0] as Error).message).to.eq(err.message);
          done();
        })
        .catch(error => {
          done(error);
        });
    });

    test('should show warning message if normal path does not exist', function (done) {
      const msg = 'A file must be open in the OOXML Viewer to search its parts.';
      const normalSubfolderPathStub = stub(cache, 'normalSubfolderPath').get(() => 'normal/path');
      const fileExistsStub = stub(FileSystemUtilities, 'fileExists').returns(Promise.resolve(false));
      const showInputStub = stub(ExtensionUtilities, 'showInput').throws(Promise.resolve('term'));
      const findInFilesStub = stub(ExtensionUtilities, 'findInFiles').returns(Promise.resolve());
      const showWarningStub = stub(ExtensionUtilities, 'showWarning').returns(Promise.resolve());
      stubs.push(normalSubfolderPathStub, showInputStub, fileExistsStub, findInFilesStub, showWarningStub);

      ooxmlPackage.searchOOXMLParts().then(() => {
        expect(showWarningStub.args[0][0]).to.eq(msg);
        done();
      });
    });
  });

  suite('openOOXMLPackage', () => {
    test('should add a files to the side bar when the OOXML package is opened', async function () {
      const packageContents = [
        {
          filePath: 'doc',
          isDirectory: true,
          data: new Uint8Array(),
        },
        {
          filePath: 'doc/document.xml',
          isDirectory: false,
          data: new TextEncoder().encode('<?xml?>'),
        },
      ];
      const fileNode = new FileNode();
      ooxmlFileAccessor.getPackageContents.returns(Promise.resolve(packageContents));
      ooxmlPackageTreeView.getRootFileNode.returns(fileNode);

      await ooxmlPackage.openOOXMLPackage();

      expect(fileNode.children.length).to.eq(1);
      expect(fileNode.children[0].fileName).to.eq('doc');
      expect(fileNode.children[0].fullPath).to.eq('');
      expect(fileNode.children[0].iconPath).to.eq(ThemeIcon.Folder);
      expect(fileNode.children[0].contextValue).to.eq('folder');
      expect(fileNode.children[0].collapsibleState).to.eq(TreeItemCollapsibleState.Expanded);
      expect(fileNode.children[0].children.length).to.eq(1);
      expect(fileNode.children[0].children[0].fileName).to.eq('document.xml');
      expect(fileNode.children[0].children[0].fullPath).to.eq('doc/document.xml');
      expect(fileNode.children[0].children[0].iconPath).to.eq(ThemeIcon.File);
      expect(fileNode.children[0].children[0].contextValue).to.eq('file');
      expect(fileNode.children[0].children[0].collapsibleState).to.eq(TreeItemCollapsibleState.None);
    });

    test('should have a file icon on first open', async function () {
      const packageContents = [
        {
          filePath: 'document.xml',
          isDirectory: false,
          data: new TextEncoder().encode('<?xml?>'),
        },
      ];
      const fileNode = new FileNode();
      ooxmlFileAccessor.getPackageContents.returns(Promise.resolve(packageContents));
      ooxmlPackageTreeView.getRootFileNode.returns(fileNode);

      await ooxmlPackage.openOOXMLPackage();

      expect(fileNode.children.length).to.eq(1);
      expect(fileNode.children[0].iconPath).to.eq(ThemeIcon.File);
    });

    test('should have a green icon if file is new on second open', async function () {
      const packageContents = [
        {
          filePath: 'document.xml',
          isDirectory: false,
          data: new TextEncoder().encode('<?xml?>'),
        },
      ];
      const fileNode = new FileNode();
      ooxmlFileAccessor.getPackageContents.onCall(0).returns(Promise.resolve([]));
      ooxmlFileAccessor.getPackageContents.onCall(1).returns(Promise.resolve(packageContents));
      ooxmlPackageTreeView.getRootFileNode.returns(fileNode);

      await ooxmlPackage.openOOXMLPackage();
      await ooxmlPackage.openOOXMLPackage();

      expect(fileNode.children.length).to.eq(1);
      expect((fileNode.children[0].iconPath as Uri).fsPath).to.contain('green');
    });

    test('should have a yellow icon if file is new on second open', async function () {
      const packageContents = [
        {
          filePath: 'document.xml',
          isDirectory: false,
          data: new TextEncoder().encode('<?xml?>'),
        },
      ];
      const fileNode = new FileNode();
      ooxmlFileAccessor.getPackageContents.onCall(0).returns(Promise.resolve(packageContents));
      ooxmlFileAccessor.getPackageContents.onCall(1).returns(Promise.resolve(packageContents));
      cache.getCachedPrevFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?><></>')));
      ooxmlPackageTreeView.getRootFileNode.returns(fileNode);

      await ooxmlPackage.openOOXMLPackage();
      await ooxmlPackage.openOOXMLPackage();

      expect(fileNode.children.length).to.eq(1);
      expect((fileNode.children[0].iconPath as Uri).fsPath).to.contain('yellow');
    });

    test('should have a red icon if file is deleted on second open', async function () {
      const packageContents = [
        {
          filePath: 'document.xml',
          isDirectory: false,
          data: new TextEncoder().encode('<?xml?>'),
        },
      ];
      const fileNode = new FileNode();
      ooxmlFileAccessor.getPackageContents.onCall(0).returns(Promise.resolve(packageContents));
      ooxmlFileAccessor.getPackageContents.onCall(1).returns(Promise.resolve([]));
      ooxmlPackageTreeView.getRootFileNode.returns(fileNode);

      await ooxmlPackage.openOOXMLPackage();
      await ooxmlPackage.openOOXMLPackage();

      expect(fileNode.children.length).to.eq(1);
      expect((fileNode.children[0].iconPath as Uri).fsPath).to.contain('red');
    });

    test('should have a green icon on third open if file is deleted on second open', async function () {
      const packageContents = [
        {
          filePath: 'document.xml',
          isDirectory: false,
          data: new TextEncoder().encode('<?xml?>'),
        },
      ];
      const fileNode = new FileNode();
      ooxmlFileAccessor.getPackageContents.onCall(0).returns(Promise.resolve(packageContents));
      ooxmlFileAccessor.getPackageContents.onCall(1).returns(Promise.resolve([]));
      ooxmlFileAccessor.getPackageContents.onCall(2).returns(Promise.resolve(packageContents));
      ooxmlPackageTreeView.getRootFileNode.returns(fileNode);

      await ooxmlPackage.openOOXMLPackage();
      await ooxmlPackage.openOOXMLPackage();
      await ooxmlPackage.openOOXMLPackage();

      expect(fileNode.children.length).to.eq(1);
      expect((fileNode.children[0].iconPath as Uri).fsPath).to.contain('green');
    });

    test('should remove file node after file is deleted on second open', async function () {
      const packageContents = [
        {
          filePath: 'document.xml',
          isDirectory: false,
          data: new TextEncoder().encode('<?xml?>'),
        },
      ];
      const fileNode = new FileNode();
      ooxmlFileAccessor.getPackageContents.onCall(0).returns(Promise.resolve(packageContents));
      ooxmlFileAccessor.getPackageContents.onCall(1).returns(Promise.resolve([]));
      ooxmlFileAccessor.getPackageContents.onCall(2).returns(Promise.resolve([]));
      ooxmlPackageTreeView.getRootFileNode.returns(fileNode);

      await ooxmlPackage.openOOXMLPackage();
      await ooxmlPackage.openOOXMLPackage();
      await ooxmlPackage.openOOXMLPackage();

      expect(fileNode.children.length).to.eq(0);
    });

    test('should showError if exception is thrown', async function () {
      const err = new Error('Error getting contents');
      ooxmlFileAccessor.getPackageContents.throws(err);
      const errorStub = stub(ExtensionUtilities, 'showError').returns(Promise.resolve());
      stubs.push(errorStub);

      await ooxmlPackage.openOOXMLPackage();

      expect(errorStub.callCount).to.eq(1);
      expect((errorStub.args[0][0] as Error).message).to.eq(err.message);
    });
  });

  suite('updateOOXMLPackage', () => {
    test('should update ooxml if there is a difference between the normal and the prev files', async function () {
      cache.cachePathIsNormal.returns(true);
      cache.getCachedNormalFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?>')));
      cache.getCachedPrevFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?><></>')));
      ooxmlFileAccessor.updatePackage.returns(Promise.resolve(true));

      await ooxmlPackage.updateOOXMLFile('file/path');

      expect(ooxmlFileAccessor.updatePackage.callCount).to.eq(1);
    });

    test('should not update ooxml if the cache path does not belong to the ', async function () {
      cache.cachePathIsNormal.returns(false);
      cache.getCachedNormalFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?>')));
      cache.getCachedPrevFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?><></>')));
      ooxmlFileAccessor.updatePackage.returns(Promise.resolve(true));

      await ooxmlPackage.updateOOXMLFile('file/path');

      expect(ooxmlFileAccessor.updatePackage.callCount).to.eq(0);
    });

    test('should not update ooxml if there is no difference between the normal and the prev files', async function () {
      cache.cachePathIsNormal.returns(true);
      cache.getCachedNormalFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?>')));
      cache.getCachedPrevFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?>')));
      ooxmlFileAccessor.updatePackage.returns(Promise.resolve(true));

      await ooxmlPackage.updateOOXMLFile('file/path');

      expect(ooxmlFileAccessor.updatePackage.callCount).to.eq(0);
    });

    test('should make editor dirty and displays warning if update package call fails', async function () {
      cache.cachePathIsNormal.returns(true);
      cache.getCachedNormalFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?>')));
      cache.getCachedPrevFile.returns(Promise.resolve(new TextEncoder().encode('<?xml?><></>')));
      ooxmlFileAccessor.updatePackage.returns(Promise.resolve(false));
      const makeDirtyStub = stub(ExtensionUtilities, 'makeActiveTextEditorDirty').returns(Promise.resolve());
      const showWarningStub = stub(ExtensionUtilities, 'showWarning').returns(Promise.resolve());
      stubs.push(makeDirtyStub, showWarningStub);
      await ooxmlPackage.updateOOXMLFile('file/path');

      expect(makeDirtyStub.callCount).to.eq(1);
      expect(showWarningStub.callCount).to.eq(1);
    });

    test('should showError if exception is thrown', async function () {
      cache.cachePathIsNormal.returns(true);
      const err = new Error('Error getting contents');
      cache.getFilePathFromCacheFilePath.throws(err);
      const errorStub = stub(ExtensionUtilities, 'showError').returns(Promise.resolve());
      stubs.push(errorStub);

      await ooxmlPackage.updateOOXMLFile('file/path');

      expect(errorStub.callCount).to.eq(1);
      expect((errorStub.args[0][0] as Error).message).to.eq(err.message);
    });
  });
});
