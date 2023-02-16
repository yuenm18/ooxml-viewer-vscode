import { expect } from 'chai';
import { spy } from 'sinon';
import { Disposable } from 'vscode';
import { OOXMLPackage } from '../../../src/ooxml-package/ooxml-package';
import { OOXMLPackageFileWatcher } from '../../../src/ooxml-package/ooxml-package-file-watcher';

suite('OOXMLPackageFileWatcher', function () {
  let fileWatchers: OOXMLPackageFileWatcher;

  setup(function () {
    fileWatchers = new OOXMLPackageFileWatcher('file-name', <OOXMLPackage>{});
  });

  test('should dispose the array of file system watchers', async function () {
    const disposeStub = spy();
    const disposable1 = {
      dispose: disposeStub,
    } as Disposable;
    const disposable2 = {
      dispose: disposeStub,
    } as Disposable;
    (fileWatchers as any).watchers.push(disposable1, disposable2);

    await fileWatchers.dispose();

    expect(disposeStub.calledTwice).to.be.true;
  });
});
