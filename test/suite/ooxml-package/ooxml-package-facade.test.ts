import { expect } from 'chai';
import { createStubInstance, SinonStubbedInstance } from 'sinon';
import { OOXMLPackage } from '../../../src/ooxml-package/ooxml-package';
import { OOXMLPackageFacade } from '../../../src/ooxml-package/ooxml-package-facade';
import { OOXMLPackageFileCache } from '../../../src/ooxml-package/ooxml-package-file-cache';
import { OOXMLPackageFileWatcher } from '../../../src/ooxml-package/ooxml-package-file-watcher';
import { OOXMLPackageTreeView } from '../../../src/ooxml-package/ooxml-package-tree-view';

suite('OOXMLPackageFacade', function () {
  let packageFacade: SinonStubbedInstance<OOXMLPackageFacade>;
  let ooxmlPackage: SinonStubbedInstance<OOXMLPackage>;
  let treeView: SinonStubbedInstance<OOXMLPackageTreeView>;
  let fileWatchers: SinonStubbedInstance<OOXMLPackageFileWatcher>;
  let fileCache: SinonStubbedInstance<OOXMLPackageFileCache>;

  setup(function () {
    ooxmlPackage = createStubInstance(OOXMLPackage);
    treeView = createStubInstance(OOXMLPackageTreeView);
    fileWatchers = createStubInstance(OOXMLPackageFileWatcher);
    fileCache = createStubInstance(OOXMLPackageFileCache);

    packageFacade = new (<any>OOXMLPackageFacade)('path', ooxmlPackage, treeView, fileWatchers, fileCache);
  });

  test('should reset file watchers, package root node, and file cache on dispose', async function () {
    await packageFacade.dispose();

    expect(fileWatchers.dispose.callCount).to.equal(1);
    expect(fileCache.reset.callCount).to.equal(1);
    expect(treeView.reset.callCount).to.equal(1);
  });
});
