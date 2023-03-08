import { workspace } from 'vscode';

/**
 * Gets the extension's settings.
 *
 * @returns {OOXMLExtensionSettings} The extension settings.
 */
export function getExtensionSettings(): OOXMLExtensionSettings {
  const ooxmlViewerConfigurationSection = workspace.getConfiguration('ooxmlViewer');
  return {
    preserveComments: ooxmlViewerConfigurationSection.get('preserveComments') ?? true,
    maximumOOXMLFileSizeBytes: ooxmlViewerConfigurationSection.get('maximumOoxmlFileSizeBytes') ?? 50000000,
    maximumNumberOfOOXMLParts: ooxmlViewerConfigurationSection.get('maximumNumberOfOoxmlParts') ?? 1000,
    maximumXmlPartsFileSizeBytes: ooxmlViewerConfigurationSection.get('maximumXmlPartsFileSizeBytes') ?? 1000000,
  };
}

/**
 * Settings for the extension.
 */
export interface OOXMLExtensionSettings {
  preserveComments: boolean;
  maximumXmlPartsFileSizeBytes: number;
  maximumNumberOfOOXMLParts: number;
  maximumOOXMLFileSizeBytes: number;
}
