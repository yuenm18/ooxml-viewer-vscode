import { expect } from 'chai';
import { dirname, join } from 'path';
import { SinonStub, stub } from 'sinon';
import { ImportMock } from 'ts-mock-imports';
import { commands, ExtensionContext, Uri } from 'vscode';
import { OOXMLTreeDataProvider } from '../../ooxml-tree-view-provider';
import { OOXMLViewer } from '../../ooxml-viewer';

suite('OOXMLViewer', function () {
  // this.timeout(10000);
  let ooxmlViewer: OOXMLViewer;
  const stubs: SinonStub[] = [];
  const testFilePath = join(__dirname, '..', '..', '..', 'test-data', 'Test.pptx');
  let context: { [key: string]: (path?: string | undefined) => string };

  setup(function () {
    context = {
      asAbsolutePath: (path?: string) => {
        return 'tacocat';
      },
    };
    ooxmlViewer = new OOXMLViewer((context as unknown) as ExtensionContext);
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
  });

  test('It should have an instance of OOXMLTreeDataProvider', function (done) {
    expect(ooxmlViewer.treeDataProvider).to.be.instanceOf(OOXMLTreeDataProvider);
    done();
  });
  test('It should populate the sidebar tree with the contents of an ooxml file', async function () {
    stubs.push(
      stub(OOXMLViewer, <never>'_mkdirPromise').returns(Promise.resolve()),
      stub(OOXMLViewer, <never>'_execPromise').returns(Promise.resolve({ stderr: null })),
      stub(OOXMLViewer, <never>'_writeFilePromise').returns(Promise.resolve()),
      stub(OOXMLViewer, <never>'_fileHasBeenChangedFromOutside').returns(Promise.resolve(false)),
    );
    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(0);
    await ooxmlViewer.viewContents(Uri.file(testFilePath));
    expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(4);
  });
  test.only('viewFile should open a text editor when called with the path to an xml file', async function () {
    const commandsMock = ImportMock.mockFunction(commands, 'executeCommand');
    stubs.push(
      stub(OOXMLViewer.prototype, <never>'_createFile').callsFake(() => Promise.resolve()),
      stub(OOXMLViewer, <never>'_readFilePromise')
        .onSecondCall()
        .returns(
          Promise.resolve(
            '<?xml version="1.0" encoding="UTF-8"?><note><to>Tove</to><from>Jani</from>' +
              "<heading>Reminder</heading><body>Don't forget me this weekend!</body></note>",
          ),
        )
        .callThrough(),
      stub(OOXMLViewer, <never>'_fileHasBeenChangedFromOutside').returns(Promise.resolve(false)),
      commandsMock,
    );
    const ooxmlViewer2 = new OOXMLViewer((context as unknown) as ExtensionContext);
    await ooxmlViewer2.viewContents(Uri.file(testFilePath));
    await ooxmlViewer2.viewFile(ooxmlViewer2.treeDataProvider.rootFileNode.children[0]);
    expect(1).to.be.gt(0);
    const folderPath = join(OOXMLViewer.fileCachePath, dirname(ooxmlViewer2.treeDataProvider.rootFileNode.children[0].fullPath));
    const filePath: string = join(folderPath, ooxmlViewer2.treeDataProvider.rootFileNode.children[0].fileName);
    expect(OOXMLViewer.openTextEditors[filePath]).to.eq(ooxmlViewer2.treeDataProvider.rootFileNode.children[0]);
    expect(commandsMock.calledWith('vscode.open')).to.be.true;
  });
  // suiteSetup(function () {
  //   // stubs.push(stub(OOXMLViewer, 'mkdirp').callsFake(function (dir: string, opts?: string | number | Options | undefined) {
  //   //   return Promise.resolve(undefined);
  //   // }));
  //   stubs.push(
  //     stub(OOXMLViewer, 'execPromise').callsFake(function (
  //       command: string,
  //       options?: (BaseEncodingOptions & ExecOptions) | null | undefined,
  //     ) {
  //       expect(command.includes('attrib +h')).to.be.true;
  //       expect(command.includes(OOXMLViewer.fileCachePath)).to.be.true;
  //       // eslint-disable-next-line @typescript-eslint/no-explicit-any
  //       return Promise.resolve({ stdout: '', stderr: '' }) as PromiseWithChild<any>;
  //     }),
  //   );
  //   stubs.push(
  //     stub(vscode.workspace, 'openTextDocument').callsFake(function (
  //       options?: { language?: string | undefined; content?: string | undefined } | undefined,
  //     ) {
  //       // eslint-disable-next-line @typescript-eslint/no-explicit-any
  //       const opts: any = options as any;
  //       expect(opts?.path.includes(OOXMLViewer.fileCachePath));
  //       return Promise.resolve('foobar') as Thenable<vscode.TextDocument>;
  //     }),
  //   );
  //   stubs.push(
  //     stub(vscode.window, 'showTextDocument').callsFake(function (uri: vscode.Uri, options?: vscode.TextDocumentShowOptions | undefined) {
  //       expect(uri).to.eq('foobar');
  //       return Promise.resolve() as Thenable<vscode.TextEditor>;
  //     }),
  //   );
  // });
  // test('It should open a ooxml part in a TextEditor', async function () {
  //   const testFilePath = join(__dirname, '..', '..', '..', 'test-data', 'Test.pptx');
  //   await ooxmlViewer.viewContents(vscode.Uri.parse(`file:///${testFilePath}`));
  //   const node: FileNode | undefined = ooxmlViewer.treeDataProvider.rootFileNode.children.find((t: FileNode) => t.fileName === 'ppt');
  //   let fileNode: FileNode | undefined;
  //   let xmlNode: FileNode | undefined;
  //   if (node?.children.length) {
  //     fileNode = node.children.find(t => t.fileName === 'slideLayouts');
  //   }
  //   if (fileNode?.children.length) {
  //     xmlNode = fileNode.children.find(t => t.fileName.endsWith('xml'));
  //   }
  //   if (xmlNode) {
  //     await ooxmlViewer.viewFile(xmlNode as FileNode);
  //   }
  //   return;
  // });
  // suiteSetup(function () {
  //   stubs.push(
  //     stub(OOXMLViewer, 'existsSync').callsFake(function (path: PathLike) {
  //       return false;
  //     }),
  //   );
  // });
  // test('It should close the editor when clear is called', function () {
  //   const testFilePath = join(__dirname, '..', '..', '..', 'test-data', 'Test.pptx');
  //   return ooxmlViewer
  //     .viewContents(vscode.Uri.parse(`file:///${testFilePath}`))
  //     .then(() => {
  //       const node: FileNode | undefined = ooxmlViewer.treeDataProvider.rootFileNode.children.find(t => t.fileName === 'ppt');
  //       let fileNode: FileNode | undefined;
  //       let xmlNode: FileNode | undefined;
  //       if (node?.children.length) {
  //         fileNode = node.children.find(t => t.fileName === 'slideLayouts');
  //       }
  //       if (fileNode?.children.length) {
  //         xmlNode = fileNode.children.find(t => t.fileName.endsWith('xml'));
  //       }
  //       return xmlNode;
  //     })
  //     .then(xmlNode => {
  //       if (xmlNode) {
  //         ooxmlViewer
  //           .viewFile(xmlNode as FileNode)
  //           .then(() => {
  //             expect(vscode.workspace.textDocuments.length).to.eq(2);
  //             expect(vscode.workspace.textDocuments[1].fileName.endsWith('xml')).to.be.true;
  //           })
  //           .then(() => {
  //             ooxmlViewer.clear().then(() => {
  //               expect(vscode.workspace.textDocuments.length).to.eq(1);
  //               expect(ooxmlViewer.treeDataProvider.rootFileNode.children.length).to.eq(0);
  //             });
  //           });
  //       }
  //     });
  // });
  suiteTeardown(function () {
    stubs.forEach(s => s.restore());
  });
});
