import { commands, Position, ProgressLocation, Range, TextEditorEdit, Uri, window, workspace } from 'vscode';
import packageJson from '../../package.json';

const extensionName = packageJson.displayName;

/**
 * Utilities for interacting with the extension host.
 */
export class ExtensionUtilities {
  /**
   * Make a text editor tab dirty.
   */
  static async makeActiveTextEditorDirty(): Promise<void> {
    window.activeTextEditor?.edit(async (textEditorEdit: TextEditorEdit) => {
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
   * Handles an error.
   *
   * @param {unknown} err The error.
   */
  static async showError(err: unknown): Promise<void> {
    let msg = 'unknown error';

    if (typeof err === 'string') {
      msg = err;
    } else if (err instanceof Error) {
      msg = err.message;
    }

    await window.showErrorMessage(msg);
  }

  /**
   * Displays a warning message.
   *
   * @param {string} message The warning message.
   */
  static async showWarning(message: string): Promise<void> {
    await window.showWarningMessage(message, { modal: true });
  }

  /**
   * Displays a warning message.
   *
   * @param {string} title The input title.
   * @param {string} prompt The input prompt.
   * @returns {Promise<string | undefined} A promise resolving to the string the user imported.
   */
  static async showInput(title: string, prompt: string): Promise<string | undefined> {
    return await window.showInputBox({ title: title, prompt: prompt });
  }

  /**
   * Runs a function while showing a progress indicator.
   *
   * @param {() => Promise<void>} func The function to call while showing the progress indicator.
   * @param {string} loadingMessage The message to show with the progress indicator.
   */
  static async withProgress(func: () => Promise<void>, loadingMessage: string): Promise<void> {
    await window.withProgress(
      {
        location: ProgressLocation.Notification,
        title: extensionName,
      },
      async progress => {
        progress.report({ message: loadingMessage });
        await func();
      },
    );
  }

  /**
   * Gets the file paths of all open text documents.
   *
   * @returns {string[]} An array of file paths.
   */
  static getOpenTextDocumentFilePaths(): string[] {
    return workspace.textDocuments.map(w => w.fileName);
  }

  /**
   * Closes an object text document.
   *
   * @param {string} fileName The file name of the text document.
   */
  static async closeTextDocument(fileName: string): Promise<void> {
    await window.showTextDocument(Uri.file(fileName), { preview: true, preserveFocus: false });
    await commands.executeCommand('workbench.action.closeActiveEditor');
  }

  /**
   * Opens a file.
   *
   * @param {string} filePath The file path.
   */
  static async openFile(filePath: string): Promise<void> {
    await commands.executeCommand('vscode.open', Uri.file(filePath));
  }

  /**
   * Runs a query against a group of files.
   *
   * @param {string} query The search query.
   * @param {string} filesToInclude The path to the files to include in the search.
   */
  static async findInFiles(query: string, filesToInclude: string): Promise<void> {
    await commands.executeCommand('workbench.action.findInFiles', {
      query: query,
      filesToInclude: filesToInclude,
      triggerSearch: true,
      isCaseSensitive: false,
      matchWholeWord: false,
    });
  }

  static async openDiff(filePath1: string, filePath2: string, title: string): Promise<void> {
    await commands.executeCommand('vscode.diff', Uri.file(filePath1), Uri.file(filePath2), title);
  }
}
