import { workspace } from 'vscode';

/**
 * Gets the extension's settings.
 *
 * @returns {OOXMLExtensionSettings} The extension settings.
 */
export function getExtensionSettings(): OOXMLExtensionSettings {
  return {
    preserveComments: workspace.getConfiguration('ooxmlViewer').get('preserveComments') ?? true,
  };
}

/**
 * Settings for the extension.
 */
export interface OOXMLExtensionSettings {
  preserveComments: boolean;
}
