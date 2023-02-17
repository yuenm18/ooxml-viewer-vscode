import { expect } from 'chai';
import { join } from 'path';
import { OOXMLPackageFileAccessor } from '../../../src/ooxml-package/ooxml-package-file-accessor';

suite('OOXMLPackageFileAccessor Integration', function () {
  const testFilePath = join(__dirname, '..', '..', '..', '..', 'test', 'test-data', 'Test.pptx');

  test('should return empty array if package is accessed before loaded', async function () {
    const fileAccessor = new OOXMLPackageFileAccessor(testFilePath);

    const contents = await fileAccessor.getPackageContents();

    expect(contents.length).to.be.eq(0);
  });

  test('should return false if package is updated before loaded', async function () {
    const fileAccessor = new OOXMLPackageFileAccessor(testFilePath);

    const response = await fileAccessor.updatePackage('file.xml', new Uint8Array());

    expect(response).to.be.eq(false);
  });

  test('should throw if loaded with invalid url', async function () {
    const fileAccessor = new OOXMLPackageFileAccessor('invalid');

    await expect(async () => await fileAccessor.load()).to.throw;
  });

  test('should load data with valid url', async function () {
    const fileAccessor = new OOXMLPackageFileAccessor(testFilePath);

    await fileAccessor.load();
  });

  test('should return package contents', async function () {
    const fileAccessor = new OOXMLPackageFileAccessor(testFilePath);
    await fileAccessor.load();

    const contents = await fileAccessor.getPackageContents();

    expect(contents).to.have.lengthOf(40);
  });
});
