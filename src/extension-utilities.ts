import { FindResult } from 'find-in-files';
import { basename } from 'path';
import { commands, Position, Range, TextDocument, TextEditor, TextEditorEdit, window } from 'vscode';

export class ExtensionUtilities {
  /**
   * @description Make a text editor tab dirty.
   * @method makeTextEditorDirty
   * @async
   * @param {TextEditor} activeTextEditor the text editor to be made dirty.
   * @returns {Promise<void>}
   */
  static async makeTextEditorDirty(activeTextEditor?: TextEditor): Promise<void> {
    activeTextEditor?.edit(async (textEditorEdit: TextEditorEdit) => {
      if (window.activeTextEditor?.selection) {
        const { activeTextEditor } = window;

        if (activeTextEditor && activeTextEditor.document.lineCount >= 2) {
          const lineNumber = activeTextEditor.document.lineCount - 2;
          const lastLineRange = new Range(new Position(lineNumber, 0), new Position(lineNumber + 1, 0));
          const lastLineText = activeTextEditor.document.getText(lastLineRange);
          textEditorEdit.replace(lastLineRange, lastLineText);
          return;
        }

        // Try to replace the first character.
        const range = new Range(new Position(0, 0), new Position(0, 1));
        const text: string | undefined = activeTextEditor?.document.getText(range);
        if (text) {
          textEditorEdit.replace(range, text);
          return;
        }

        // With an empty file, we first add a character and then remove it.
        // This has to be done as two edits, which can cause the cursor to
        // visibly move and then return, but we can at least combine them
        // into a single undo step.
        await activeTextEditor?.edit(
          (innerEditBuilder: TextEditorEdit) => {
            innerEditBuilder.replace(range, ' ');
          },
          { undoStopBefore: true, undoStopAfter: false },
        );

        await activeTextEditor?.edit(
          (innerEditBuilder: TextEditorEdit) => {
            innerEditBuilder.replace(range, '');
          },
          { undoStopBefore: false, undoStopAfter: true },
        );
      }
    });
  }

  /**
   * @description Closes text editors.
   * @method closeEditors
   * @async
   * @param {TextDocument[]} textDocuments The text documents to close.
   * @returns {Promise<void>}
   */
  static async closeEditors(textDocuments: TextDocument[]): Promise<void> {
    try {
      if (textDocuments.length) {
        const td: TextDocument | undefined = textDocuments.pop();
        if (td) {
          await window.showTextDocument(td, { preview: true, preserveFocus: false });
          await commands.executeCommand('workbench.action.closeActiveEditor');
          await this.closeEditors(textDocuments);
        }
      }
    } catch (err) {
      console.error(err);
    }
  }

  static generateHtml(results: FindResult): string {
    let searchResultHtml = '<div class="list-group">';

    Object.keys(results).forEach((key: string) => {
      searchResultHtml += `<button
                              id="${key}"
                              type="button"
                              class="list-group-item list-group-item-action d-flex justify-content-between align-items-start"
                              onclick="openPart(this.id)"
                            >
                            <div class="ms-2 me-auto">
                              <div class="fw-bold">${basename(key)}</div>
                            </div>
                            <span class="badge bg-primary rounded-pill text-white">${results[key].count}</span>
                          </li>`;
    });

    searchResultHtml += '</div>';

    return `<!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>
          <title>OOXML Validation Errors</title>
          <body>
            <div class="container">
              <div class="row">
                <div class="col">
                  ${searchResultHtml}
                </div>
              </div>
                </div>
              </div>
            </div>
            <script>
            const vscode = acquireVsCodeApi();

            function openPart (id) {
              vscode.postMessage({
                       command: 'openPart',
                       text: id,
                     });
            }
          </script>
          </body>
        </html>`;
  }
}
