import { expect } from 'chai';
import { tmpdir } from 'os';
import { join } from 'path';
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
    const writeFileStub = stub(ooxmlFileCache, <never>'writeFile').returns(Promise.resolve());
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
      const writeFileStub = stub(ooxmlFileCache, <never>'writeFile').returns(Promise.resolve());
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
      const readFileStub = stub(workspace.fs, 'readFile').returns(Promise.resolve(oldFileContents));
      const writeFileStub = stub(workspace.fs, 'writeFile').returns(Promise.resolve());
      const createDirectoryStub = stub(workspace.fs, 'createDirectory').returns(Promise.resolve());
      stubs.push(readFileStub, writeFileStub, createDirectoryStub);

      await ooxmlFileCache.updateCachedFile(filePath, fileContents, true);

      expect(writeFileStub.callCount).to.equal(3);
      expect(createDirectoryStub.callCount).to.equal(3);
      expect(writeFileStub.calledWith(match(fileCacheUri), fileContents)).to.be.true;
      expect(writeFileStub.calledWith(match(prevFileCacheUri), fileContents)).to.be.true;
      expect(writeFileStub.calledWith(match(compareFileCacheUri), oldFileContents)).to.be.true;
    },
  );

  test(
    'should update cachedFile, prevCachedFile with contents' + ' when updateCachedFile is called with updateCompareFile=false',
    async function () {
      const fileContents = new TextEncoder().encode('new content');
      const oldFileContents = new TextEncoder().encode('old content');
      const readFileStub = stub(workspace.fs, 'readFile').returns(Promise.resolve(oldFileContents));
      const writeFileStub = stub(workspace.fs, 'writeFile').returns(Promise.resolve());
      const createDirectoryStub = stub(workspace.fs, 'createDirectory').returns(Promise.resolve());
      stubs.push(readFileStub, writeFileStub, createDirectoryStub);

      await ooxmlFileCache.updateCachedFile(filePath, fileContents, false);

      expect(writeFileStub.callCount).to.equal(2);
      expect(createDirectoryStub.callCount).to.equal(2);
      expect(writeFileStub.calledWith(match(fileCacheUri), fileContents)).to.be.true;
      expect(writeFileStub.calledWith(match(prevFileCacheUri), fileContents)).to.be.true;
      expect(writeFileStub.calledWith(match(compareFileCacheUri), oldFileContents)).to.be.false;
    },
  );

  test('should update compareFile with contents when updateCompareFile is called', async function () {
    const fileContents = new TextEncoder().encode('test');
    const writeFileStub = stub(workspace.fs, 'writeFile').returns(Promise.resolve());
    const createDirectoryStub = stub(workspace.fs, 'createDirectory').returns(Promise.resolve());
    stubs.push(writeFileStub, createDirectoryStub);

    await ooxmlFileCache.updateCompareFile(filePath, fileContents);

    expect(writeFileStub.callCount).to.equal(1);
    expect(createDirectoryStub.callCount).to.equal(1);
    expect(writeFileStub.calledWith(match(compareFileCacheUri), fileContents)).to.be.true;
  });

  test('should deleted cachedFile, prevCachedFile and compareCachedFile when deleteCachedFile is called', async function () {
    const deleteFileStub = stub(workspace.fs, 'delete').returns(Promise.resolve());
    stubs.push(deleteFileStub);

    await ooxmlFileCache.deleteCachedFiles(filePath);

    expect(deleteFileStub.callCount).to.equal(3);
    expect(deleteFileStub.calledWith(match(fileCacheUri))).to.be.true;
    expect(deleteFileStub.calledWith(match(prevFileCacheUri))).to.be.true;
    expect(deleteFileStub.calledWith(match(compareFileCacheUri))).to.be.true;
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
    const readFileStub = stub(workspace.fs, 'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect(readFileStub.calledWith(match(fileCacheUri))).to.be.true;
    expect(result).to.be.equal(fileContents);
  });

  test('should get cached file when getCachedPrevFile is called', async function () {
    const fileContents = new TextEncoder().encode('text');
    const readFileStub = stub(workspace.fs, 'readFile').returns(Promise.resolve(fileContents));
    stubs.push(readFileStub);

    const result = await ooxmlFileCache.getCachedPrevFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect(readFileStub.calledWith(match(prevFileCacheUri))).to.be.true;
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

  test('should log error and return empty array if readFile fails', async function () {
    const readFileStub = stub(workspace.fs, 'readFile').throwsException();
    const errorLogStub = stub(console, 'error');
    stubs.push(readFileStub, errorLogStub);

    const result = await ooxmlFileCache.getCachedFile(filePath);

    expect(readFileStub.callCount).to.equal(1);
    expect(readFileStub.calledWith(match(fileCacheUri))).to.be.true;
    expect(result).to.be.eql(new Uint8Array());
    expect(errorLogStub.called).to.be.true;
  });

  test('should log error if creating file fails', async function () {
    const readFileStub = stub(workspace.fs, 'writeFile').throwsException();
    const errorLogStub = stub(console, 'error');
    stubs.push(readFileStub, errorLogStub);

    await ooxmlFileCache.createCachedFile(filePath, new Uint8Array(), false);

    expect(errorLogStub.called).to.be.true;
  });

  test('should log error if delete fails', async function () {
    const readFileStub = stub(workspace.fs, 'delete').throwsException();
    const errorLogStub = stub(console, 'error');
    stubs.push(readFileStub, errorLogStub);

    await ooxmlFileCache.deleteCachedFiles(filePath);

    expect(errorLogStub.called).to.be.true;
  });
});
