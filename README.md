# OOXML Viewer VSCode Extension

[![Continuous Integration](https://github.com/yuenm18/ooxml-viewer-vscode/actions/workflows/ci.yaml/badge.svg)](https://github.com/yuenm18/ooxml-viewer-vscode/actions/workflows/ci.yaml)
[![Visual Studio Marketplace](https://vsmarketplacebadge.apphb.com/version/yuenm18.ooxml-viewer.svg)](https://marketplace.visualstudio.com/items?itemName=yuenm18.ooxml-viewer)
[![Open in Remote - Containers](https://img.shields.io/static/v1?label=Remote%20-%20Containers&message=Open&color=blue&logo=visualstudiocode)](https://vscode.dev/redirect?url=vscode://ms-vscode-remote.remote-containers/cloneInVolume?url=https://github.com/yuenm18/ooxml-viewer-vscode)

\* Please note, files must be stored locally, i.e. not in OneNote, Dropbox, etc.

## Features

- [Display the contents of OOXML documents in VS Code](#displays-the-contents-of-ooxml-documents-in-vs-code)
- [Edit the contents of an OOXML documents in VS Code](#edit-the-contents-of-an-ooxml-documents-in-vs-code)
- [Get diff when OOXML documents are edited from outside, e.g. in Microsoft Word, Libre Office Writer, Microsoft Excel, Libre Office Calc, etc.](#user-content-get-diff-when-ooxml-documents-are-edited-from-outside-eg-in-microsoft-word-libre-office-writer-microsoft-excel-libre-office-calc-etc)
- [Search all parts](#search-all-parts)
- [Search parts of any file that uses the Open Packaging Conventions](#search-parts-of-any-file-that-uses-the-open-packaging-conventions)

### Display the contents of OOXML documents in VS Code

To view the contents of an OOXML document, right click on the file in the context menu, then click on the part you want to view.

![Opening an OOXML Part](https://raw.githubusercontent.com/yuenm18/ooxml-viewer-vscode/master/resources/images/view-part.gif)

### Edit the contents of OOXML documents in VS Code

To edit an OOXML part, select the part in the OOXML Viewer menu then edit and save. The changes will be reflected in the OOXML document.

![Editing the contents of an OOXML document in VS Code](https://raw.githubusercontent.com/yuenm18/ooxml-viewer-vscode/master/resources/images/edit-part.gif)

### Get diff when OOXML documents are edited from outside, e.g. in Microsoft Word, Libre Office Writer, Microsoft Excel, Libre Office Calc, etc.

When a document opened by the OOXML Viewer is edited from an external program, changed parts are marked with a yellow asterisk, deleted parts are marked with a red asterisk, and new parts are marked with a green asterisk.

To view a diff with the previous version of an OOXML part, right click on the part in the OOXML Viewer menu and click "Compare with Previous".

#### Diff adding text to a odt file

![Getting the diff of an OOXML Part](https://raw.githubusercontent.com/yuenm18/ooxml-viewer-vscode/master/resources/images/edit-file.gif)

#### Diff adding a slide to a pptx file

![Getting the diff of an OOXML Part](https://raw.githubusercontent.com/yuenm18/ooxml-viewer-vscode/master/resources/images/edit-file-add.gif)

#### Diff removing a slide from a pptx file

![Getting the diff of an OOXML Part](https://raw.githubusercontent.com/yuenm18/ooxml-viewer-vscode/master/resources/images/edit-file-remove.gif)

### Search all parts

To search all parts, click "SEARCH PARTS" in the tree view title bar, enter your search term, and press enter/return. The initial search is not case sensitive or whole words only, but once the OOXML Viewer opens the search pane, all VS Code search options are available.

![Searching all OOXML Parts](https://raw.githubusercontent.com/yuenm18/ooxml-viewer-vscode/master/resources/images/find-in-parts.gif)

### Search parts of any file that uses the Open Packaging Conventions

By default, the OOXML Viewer can view and edit the contents of files with these extensions: ".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp", ".docm", ".dotm", ".xlsm", ".pptm", ".dotx", ".xltx", ".xltm", ".potx", ".sldx", ".ppsx". But the OOXML Viewer extension can be used with any file type that uses the [Open Packaging Conventions](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview) or any zip based file type.

To add additional file types, open the `settings.json` and add or update the `files.associations` and add an associate with "ooxml". For example, to add "\*.vsix" files, settings.json should include

```json
"files.associations": {
  "*.vsix": "ooxml"
}
```

After adding the file extension, restart VS Code and right click on the file to open and select "Open OOXML File" to view and edit it's contents.

## Extension Settings

This extension contributes the following variables to the [settings](https://code.visualstudio.com/docs/customization/userandworkspace):

- `ooxmlViewer.preserveComments`: Boolean, determines if comments will be preserved in XML or removed on save. Defaults to true.

## Release Notes

Please see the [Changelog](CHANGELOG.md)
