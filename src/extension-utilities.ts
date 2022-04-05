import { commands, env, FileSystemError, Position, Range, TextDocument, TextEditor, TextEditorEdit, window } from 'vscode';

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
      this.handleError(err);
    }
  }

  static async closeEditorsOnStartup(cacheBasePath: string): Promise<void> {
    /**
     * This nonsense is necessary to close all open OOXML Viewer files on startup, because of this issue with VS Code:
     * https://github.com/microsoft/vscode/issues/15178
     * on startup, VS Code only recognizes the open tab in workspace.textDocuments, giving it a length of 1, so
     * ExtensionUtilities.closeEditors(workspace.textDocuments) only closes one text editor, so we loop through all the text
     * editors to make them active and close each active editor.
     *
     * And we use the clipboard to get the file name, because window.activeTextEditor.document is undefined with
     * unsupported or binary files, so we can't use window.activeTextEditor.document.fileName
     * https://github.com/Microsoft/vscode/issues/2582#issuecomment-246692860
     *
     */
    let fileName: string | undefined;
    const fileNames: string[] = [];
    await commands.executeCommand('workbench.action.files.copyPathOfActiveFile');
    fileName = await env.clipboard.readText();

    while (!fileNames.includes(fileName)) {
      fileNames.push(fileName);

      await commands.executeCommand('workbench.action.nextEditor');

      await commands.executeCommand('workbench.action.files.copyPathOfActiveFile');
      fileName = await env.clipboard.readText();

      if (fileName.toLowerCase().startsWith(cacheBasePath.toLowerCase())) {
        await commands.executeCommand('workbench.action.closeActiveEditor');
      }
    }
  }

  static async handleError(err: unknown): Promise<void> {
    try {
      let msg = 'unknown error';

      if (typeof err === 'string') {
        msg = err;
      } else if (err instanceof Error) {
        msg = err.message;
      }

      await window.showErrorMessage(msg);
    } catch(error) {
      const msg = (error as Error)?.message || (error as FileSystemError)?.code || 'unknown error';
      throw new Error(msg);
    }
  }
}
