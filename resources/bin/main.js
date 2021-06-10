// eslint-disable-next-line no-undef
const vscode = acquireVsCodeApi();

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function openPart(id) {
  vscode.postMessage({
    command: 'openPart',
    text: id,
  });
}
