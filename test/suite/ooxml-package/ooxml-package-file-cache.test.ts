import { expect } from 'chai';
import { tmpdir } from 'os';
import { basename, join } from 'path';
import { SinonStub, stub } from 'sinon';
import { Uri, workspace } from 'vscode';
import { OOXMLPackageFileCache } from '../../../src/ooxml-package/ooxml-package-file-cache';
import { ExtensionUtilities } from '../../../src/utilities/extension-utilities';
import { FileSystemUtilities } from '../../../src/utilities/file-system-utilities';

suite('OOXMLPackageFileCache', function () {
  const stubs: SinonStub[] = [];
  let ooxmlFileCache: OOXMLPackageFileCache;

  const filePath = 'doc/document.xml';
  const cacheBasePath = join(tmpdir(), 'ooxml-viewer');
  const fileCachePath = join(cacheBasePath, 'cache', 'file-name', 'normal', 'doc/document.xml');
  const prevFileCachePath = join(cacheBasePath, 'cache', 'file-name', 'prev', 'doc/document.xml');
  const compareFileCachePath = join(cacheBasePath, 'cache', 'file-name', 'compare', 'doc/document.xml');
  const fileCacheUri = Uri.file(fileCachePath);
  const prevFileCacheUri = Uri.file(prevFileCachePath);
  const compareFileCacheUri = Uri.file(compareFileCachePath);

  setup(async function () {
    ooxmlFileCache = new OOXMLPackageFileCache('file-name', join(tmpdir(), 'ooxml-viewer'));

    await ooxmlFileCache.initialize();
  });

  teardown(async function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
    await ooxmlFileCache.reset();
  });

  test('should create cachedFile, prevCachedFile, and compareCachedFile when createCachedFiles is called', async function () {
    const writeFileStub = stub(FileSystemUtilities, 'writeFile').returns(Promise.resolve(true));
    stubs.push(writeFileStub);
    const fileContents = new TextEncoder().encode('test');

    await ooxmlFileCache.createCachedFiles(filePath, fileContents);

    expect(writeFileStub.callCount).to.equal(3);
    expect((writeFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
    expect(writeFileStub.args[0][1]).to.eq(fileContents);
    expect((writeFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
    expect(writeFileStub.args[1][1]).to.eq(fileContents);
    expect((writeFileStub.args[2][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
    expect(writeFileStub.args[2][1]).to.eq(fileContents);
  });

  test(
    'should create cachedFile and prevCachedFile with fileContents' +
      ' and compareCachedFile with empty contents when createCachedFilesWithEmptyCompare is called',
    async function () {
      const writeFileStub = stub(FileSystemUtilities, 'writeFile').returns(Promise.resolve(true));
      stubs.push(writeFileStub);
      const fileContents = new TextEncoder().encode('test');

      await ooxmlFileCache.createCachedFilesWithEmptyCompare(filePath, fileContents);

      expect(writeFileStub.callCount).to.equal(3);
      expect((writeFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[0][1]).to.eq(fileContents);
      expect((writeFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[1][1]).to.eq(fileContents);
      expect((writeFileStub.args[2][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[2][1]).to.deep.eq(new Uint8Array());
    },
  );

  test('should update cachedFile, prevCachedFile with contents when updateCachedFilesNoCompare is called', async function () {
    const fileContents = new TextEncoder().encode('new content');
    const oldFileContents = new TextEncoder().encode('old content');
    const readFileStub = stub(ooxmlFileCache, <never>'readFile').returns(Promise.resolve(oldFileContents));
    const writeFileStub = stub(ooxmlFileCache, <never>'writeFile').returns(Promise.resolve());
    stubs.push(readFileStub, writeFileStub);

    await ooxmlFileCache.updateCachedFilesNoCompare(filePath, fileContents);

    expect(writeFileStub.callCount).to.equal(2);
    expect((writeFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
    expect(writeFileStub.args[0][1]).to.eq(fileContents);
    expect((writeFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
    expect(writeFileStub.args[1][1]).to.eq(fileContents);
    expect(writeFileStub.args[2]).to.be.undefined;
  });

  test(
    'should update cachedFile, prevCachedFile with contents and compareCachedFile with cachedFile' + ' when updateCachedFiles is called',
    async function () {
      const fileContents = new TextEncoder().encode('new content');
      const oldFileContents = new TextEncoder().encode('old content');
      const readFileStub = stub(FileSystemUtilities, 'readFile').returns(Promise.resolve(oldFileContents));
      const writeFileStub = stub(FileSystemUtilities, 'writeFile').returns(Promise.resolve(true));
      stubs.push(readFileStub, writeFileStub);

      await ooxmlFileCache.updateCachedFiles(filePath, fileContents);

      expect(writeFileStub.callCount).to.equal(3);
      expect((writeFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[0][1]).to.eq(fileContents);
      expect((writeFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[1][1]).to.eq(fileContents);
      expect((writeFileStub.args[2][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[2][1]).to.deep.eq(oldFileContents);
    },
  );

  test('should update compareFile with contents when updateCompareFile is called', async function () {
    const fileContents = new TextEncoder().encode('test');
    const writeFileStub = stub(ooxmlFileCache, <never>'writeFile').returns(Promise.resolve());
    stubs.push(writeFileStub);

    await ooxmlFileCache.updateCompareFile(filePath, fileContents);

    expect(writeFileStub.callCount).to.equal(1);
    expect(basename(writeFileStub.args[0][0] as string)).to.eq(basename(filePath));
    expect(writeFileStub.args[0][1]).to.deep.eq(fileContents);
  });

  test('should deleted cachedFile, prevCachedFile and compareCachedFile when deleteNormalCachedFile is called', async function () {
    const deleteFileStub = stub(FileSystemUtilities, 'deleteFile').returns(Promise.resolve());
    stubs.push(deleteFileStub);

    await ooxmlFileCache.deleteCachedFiles(filePath);

    expect(deleteFileStub.callCount).to.equal(3);
    expect((deleteFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
    expect((deleteFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
    expect((deleteFileStub.args[2][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
  });

  test('should return file cache path when getNormalFileCachePath is called', function () {
    const fileCachePath = ooxmlFileCache.getNormalFileCachePath(filePath);

    expect(fileCachePath).to.equal(fileCachePath);
  });

  test('should return file cache path when getCompareFileCachePath is called', function () {
    const fileCachePath = ooxmlFileCache.getCompareFileCachePath(filePath);

    expect(fileCachePath).to.equal(compareFileCachePath);
  });

  test('should return file path when getFilePathFromCacheFilePath is called with a cache path', function () {
    const expectedFilePath = ooxmlFileCache.getFilePathFromCacheFilePath(fileCachePath);

    expect(expectedFilePath).to.equal(filePath);
  });

  test('should return original path when getFilePathFromCacheFilePath is called with a path not in the cache', function () {
    const notInCache = 'notincache';

    const originalPath = ooxmlFileCache.getFilePathFromCacheFilePath(notInCache);

    expect(originalPath).to.equal(notInCache);
  });

  test('should return true when pathBelongsToCache is called with a path in the cache', function () {
    const pathInCache = fileCachePath;

    const result = ooxmlFileCache.pathBelongsToCache(pathInCache);

    expect(result).to.be.true;
  });

  test('should return false when pathBelongsToCache is called with a path not in the cache', function () {
    const pathNotInCache = 'not in cache';

    const result = ooxmlFileCache.pathBelongsToCache(pathNotInCache);

    expect(result).to.be.false;
  });

  test('should return true when cachePathIsNormal is called with a path in the normal cache', function () {
    const normalCachePath = fileCachePath;

    const result = ooxmlFileCache.cachePathIsNormal(normalCachePath);

    expect(result).to.be.true;
  });

  test('should return false when cachePathIsNormal is called with a path not in the normal cache', function () {
    const pathNotInCache = prevFileCachePath;

    const result = ooxmlFileCache.cachePathIsNormal(pathNotInCache);

    expect(result).to.be.false;
  });

  test('should get cached file when getCachedNormalFile is called', async function () {
    const fileContents = new TextEncoder().encode('text');
    const readFileStub = stub(ooxmlFileCache, <never>'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedNormalFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect((readFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
    expect(result).to.deep.equal(fileContents);
  });

  test('should get cached file when getCachedPrevFile is called', async function () {
    const fileContents = new TextEncoder().encode('text');
    const readFileStub = stub(FileSystemUtilities, 'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedPrevFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect((readFileStub.args[0][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
    expect(result).to.be.equal(fileContents);
  });

  test('should get cached file when getCachedCompareFile is called', async function () {
    const fileContents = new TextEncoder().encode('text');
    const readFileStub = stub(FileSystemUtilities, 'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedCompareFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect((readFileStub.args[0][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
    expect(result).to.be.equal(fileContents);
  });

  test('should reset the cache when reset is called when delete cacheBasePath does not exist', async function () {
    workspace.fs.delete(Uri.file(cacheBasePath), { useTrash: false, recursive: true });
    const showErrorStub = stub(ExtensionUtilities, 'showError').returns(Promise.resolve(undefined));
    stubs.push(showErrorStub);

    expect(showErrorStub.callCount).to.equal(0);
    expect(async () => await workspace.fs.stat(Uri.file(cacheBasePath))).to.not.throw;
  });
});
