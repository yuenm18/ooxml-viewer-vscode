import { basename, dirname } from 'path';
import { Disposable, FileSystemWatcher, RelativePattern, workspace } from 'vscode';
import logger from '../utilities/logger';
import { OOXMLPackage } from './ooxml-package';

/**
 * The OOXML Package File Watcher.
 */
export class OOXMLPackageFileWatcher {
  private watchers: Disposable[] = [];

  /**
   * Constructs an instance of OOXMLPackageFileWatcher.
   *
   * @constructor
   * @param {string} filePath The path to the ooxml file.
   * @param {string} ooxmlPackage The ooxml package.
   */
  constructor(filePath: string, ooxmlPackage: OOXMLPackage) {
    this.setupFileWatchers(filePath, ooxmlPackage);
  }

  /**
   * Disposes the file watchers.
   */
  dispose(): void {
    this.watchers.forEach(w => w.dispose());
    this.watchers = [];
  }

  private setupFileWatchers(filePath: string, ooxmlPackage: OOXMLPackage) {
    // set up watchers
    const fileSystemWatcher: FileSystemWatcher = workspace.createFileSystemWatcher(
      new RelativePattern(dirname(filePath), basename(filePath)),
    );

    // Prevent multiple comparison operations on large files
    let locked = false;
    // Prevents multiple file system watcher triggers from unmodified files
    let stats = { mtime: 0 };

    fileSystemWatcher.onDidChange(async uri => {
      logger.trace('File system did change triggered');
      const newStats = await workspace.fs.stat(uri);

      if (!locked && stats.mtime !== newStats.mtime) {
        logger.trace('Locking file system watcher');
        locked = true;

        await ooxmlPackage.openOOXMLPackage();

        locked = false;
        stats = await workspace.fs.stat(uri);

        logger.trace('Unlocking file system watcher');
      } else {
        logger.debug('File system watcher locked');
      }
    });

    fileSystemWatcher.onDidDelete(async uri => {
      logger.trace('File system did delete triggered');

      await ooxmlPackage.removePackage();
    });

    const openTextDocumentWatcher = workspace.onDidOpenTextDocument(async document => ooxmlPackage.tryFormatDocument(document.fileName));
    const saveTextDocumentWatcher = workspace.onDidSaveTextDocument(async document => ooxmlPackage.updateOOXMLFile(document.fileName));

    this.watchers.push(openTextDocumentWatcher, fileSystemWatcher, saveTextDocumentWatcher);
  }
}
