import { LogOutputChannel, window } from 'vscode';
import packageJson from '../../package.json';
const extensionName = packageJson.displayName;

export class Logger {
  private logger: LogOutputChannel;

  constructor() {
    this.logger = window.createOutputChannel(extensionName, { log: true });
  }

  trace(message: string, ...args: any[]): void {
    this.logger.trace(message, ...args);
  }

  debug(message: string, ...args: any[]): void {
    this.logger.debug(message, ...args);
  }

  info(message: string, ...args: any[]): void {
    this.logger.info(message, ...args);
  }

  warn(message: string, ...args: any[]): void {
    this.logger.warn(message, ...args);
  }

  error(error: string | Error, ...args: any[]): void {
    this.logger.error(error, ...args);
  }
}

export default new Logger();
