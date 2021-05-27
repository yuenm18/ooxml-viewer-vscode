import { expect } from 'chai';
import { tmpdir } from 'os';
import { basename, join } from 'path';
import { match, SinonStub, stub } from 'sinon';
import { ExtensionContext, Uri, workspace } from 'vscode';
import { OOXMLFileCache } from '../../src/ooxml-file-cache';

suite('OOXMLViewer File Cache', function () {
  const stubs: SinonStub[] = [];
  let ooxmlFileCache: OOXMLFileCache;

  const filePath = 'doc/document.xml';
  const fileCachePath = join(tmpdir(), 'ooxml-viewer', 'cache', 'normal', 'doc/document.xml');
  const prevFileCachePath = join(tmpdir(), 'ooxml-viewer', 'cache', 'prev', 'doc/document.xml');
  const compareFileCachePath = join(tmpdir(), 'ooxml-viewer', 'cache', 'compare', 'doc/document.xml');
  const fileCacheUri = Uri.file(fileCachePath);
  const prevFileCacheUri = Uri.file(prevFileCachePath);
  const compareFileCacheUri = Uri.file(compareFileCachePath);

  setup(function () {
    const context = {
      storageUri: {
        fsPath: join(tmpdir(), 'ooxml-viewer'),
      },
    } as ExtensionContext;

    ooxmlFileCache = new OOXMLFileCache(context);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('should create cachedFile, prevCachedFile, and compareCachedFile when createCachedFile is called', async function () {
    const writeFileStub = stub(ooxmlFileCache, 'writeFile').returns(Promise.resolve());
    stubs.push(writeFileStub);
    const fileContents = new TextEncoder().encode('test');

    await ooxmlFileCache.createCachedFile(filePath, fileContents, false);

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
      ' and compareCachedFile with empty contents when createCachedFile is called with createEmptyCompareFile=true',
    async function () {
      const writeFileStub = stub(ooxmlFileCache, 'writeFile').returns(Promise.resolve());
      stubs.push(writeFileStub);
      const fileContents = new TextEncoder().encode('test');

      await ooxmlFileCache.createCachedFile(filePath, fileContents, true);

      expect(writeFileStub.callCount).to.equal(3);
      expect((writeFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[0][1]).to.eq(fileContents);
      expect((writeFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[1][1]).to.eq(fileContents);
      expect((writeFileStub.args[2][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[2][1]).to.deep.eq(new Uint8Array());
    },
  );

  test(
    'should update cachedFile, prevCachedFile with contents and compareCachedFile with cachedFile' +
      ' when updateCachedFile is called with updateCompareFile=true',
    async function () {
      const fileContents = new TextEncoder().encode('new content');
      const oldFileContents = new TextEncoder().encode('old content');
      const readFileStub = stub(ooxmlFileCache, 'readFile').returns(Promise.resolve(oldFileContents));
      const writeFileStub = stub(ooxmlFileCache, 'writeFile').returns(Promise.resolve());
      stubs.push(readFileStub, writeFileStub);

      await ooxmlFileCache.updateCachedFile(filePath, fileContents, true);

      expect(writeFileStub.callCount).to.equal(3);
      expect((writeFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[0][1]).to.eq(fileContents);
      expect((writeFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[1][1]).to.eq(fileContents);
      expect((writeFileStub.args[2][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
      expect(writeFileStub.args[2][1]).to.deep.eq(oldFileContents);
    },
  );

  // eslint-disable-next-line max-len
  test('should update cachedFile, prevCachedFile with contents when updateCachedFile is called with updateCompareFile=false', async function () {
    const fileContents = new TextEncoder().encode('new content');
    const oldFileContents = new TextEncoder().encode('old content');
    const readFileStub = stub(ooxmlFileCache, 'readFile').returns(Promise.resolve(oldFileContents));
    const writeFileStub = stub(ooxmlFileCache, 'writeFile').returns(Promise.resolve());
    stubs.push(readFileStub, writeFileStub);

    await ooxmlFileCache.updateCachedFile(filePath, fileContents, false);

    expect(writeFileStub.callCount).to.equal(2);
    expect((writeFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
    expect(writeFileStub.args[0][1]).to.eq(fileContents);
    expect((writeFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
    expect(writeFileStub.args[1][1]).to.eq(fileContents);
    expect(writeFileStub.args[2]).to.be.undefined;
  });

  test('should update compareFile with contents when updateCompareFile is called', async function () {
    const fileContents = new TextEncoder().encode('test');
    const writeFileStub = stub(ooxmlFileCache, 'writeFile').returns(Promise.resolve());
    stubs.push(writeFileStub);

    await ooxmlFileCache.updateCompareFile(filePath, fileContents);

    expect(writeFileStub.callCount).to.equal(1);
    expect(basename(writeFileStub.args[0][0] as string)).to.deep.eq(basename(filePath));
    expect(writeFileStub.args[0][1]).to.deep.eq(fileContents);
  });

  test('should deleted cachedFile, prevCachedFile and compareCachedFile when deleteCachedFile is called', async function () {
    const deleteFileStub = stub(ooxmlFileCache, <never>'deleteFile').returns(Promise.resolve());
    stubs.push(deleteFileStub);

    await ooxmlFileCache.deleteCachedFiles(filePath);

    expect(deleteFileStub.callCount).to.equal(3);
    expect((deleteFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
    expect((deleteFileStub.args[1][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
    expect((deleteFileStub.args[2][0] as string).toLowerCase()).to.eq(compareFileCacheUri.fsPath.toLowerCase());
  });

  test('should return file cache path when getFileCachePath is called', function () {
    const fileCachePath = ooxmlFileCache.getFileCachePath(filePath);

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

  test('should get cached file when getCachedFile is called', async function () {
    const fileContents = new TextEncoder().encode('text');
    const readFileStub = stub(ooxmlFileCache, 'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect((readFileStub.args[0][0] as string).toLowerCase()).to.eq(fileCacheUri.fsPath.toLowerCase());
    expect(result).to.deep.equal(fileContents);
  });

  test('should get cached file when getCachedPrevFile is called', async function () {
    const fileContents = new TextEncoder().encode('text');
    const readFileStub = stub(ooxmlFileCache, 'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedPrevFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect((readFileStub.args[0][0] as string).toLowerCase()).to.eq(prevFileCacheUri.fsPath.toLowerCase());
    expect(result).to.be.equal(fileContents);
  });

  test('should get cached file when getCachedCompareFile is called', async function () {
    const fileContents = new TextEncoder().encode('text');
    const readFileStub = stub(workspace.fs, 'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedCompareFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect(readFileStub.calledWith(match(compareFileCacheUri))).to.be.true;
    expect(result).to.be.equal(fileContents);
  });
});
