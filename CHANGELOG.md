# Changelog

## [Unreleased]

## [2.0.0] - 2023-06-07

### Added

- Support for multiple ooxml packages in the tree view [#34](https://github.com/yuenm18/ooxml-viewer-vscode/issues/34) [#40](https://github.com/yuenm18/ooxml-viewer-vscode/issues/40)
- Logging in output window [#43](https://github.com/yuenm18/ooxml-viewer-vscode/issues/43)

## [1.3.1] - 2023-01-16

### Fixed

- Broken badge in README and update dependencies

## [1.3.0] - 2022-08-08

### Added

- Setting to preserve comments in xml parts [#37](https://github.com/yuenm18/ooxml-viewer-vscode/issues/37)

## [1.2.2] - 2022-04-14

### Fixed

- Error showing up when opening ooxml file in a new workspace for the first time [#33](https://github.com/yuenm18/ooxml-viewer-vscode/issues/33)

## [1.2.1] - 2022-04-09

### Fixed

- Change tracking not working sometimes with large files [#27](https://github.com/yuenm18/ooxml-viewer-vscode/issues/27)
- Issue where "Compare with Previous" shows up on folders that could not be diffed [#31](https://github.com/yuenm18/ooxml-viewer-vscode/issues/31)
- "Unable to open" error when opening ooxml files when parts are already open in editor [#32](https://github.com/yuenm18/ooxml-viewer-vscode/issues/32)

## [1.2.0] - 2021-08-03

### Added

- Feature to search the whole package [#17](https://github.com/yuenm18/ooxml-viewer-vscode/issues/17)

### Fixed

- Error when opening ooxml files with command [#23](https://github.com/yuenm18/ooxml-viewer-vscode/issues/23)
- Issue where only first file would close when opening vscode [#24](https://github.com/yuenm18/ooxml-viewer-vscode/issues/24)

### Changed

- Show file name when viewing ooxml file [#22](https://github.com/yuenm18/ooxml-viewer-vscode/issues/22)

## [1.1.1] - 2021-05-28

### Fixed

- Xml spacing issue with saved documents that resulted in Word not being able to open documents with shapes or 3D models [#20](https://github.com/yuenm18/ooxml-viewer-vscode/issues/20)

### Changed

- Add DEFLATE compression to ooxml packages

## [1.1.0] - 2021-02-25

### Added

- Support for _.dotm_ files [#16](https://github.com/yuenm18/ooxml-viewer-vscode/issues/16)

### Fixed

- Xml formatting issue [#14](https://github.com/yuenm18/ooxml-viewer-vscode/issues/14)

## [1.0.2] - 2021-02-17

### Fixed

- Properly save xml parts that contain unicode characters [#15](https://github.com/yuenm18/ooxml-viewer-vscode/issues/15)

## [1.0.1] - 2021-01-07

### Fixed

- Xml file is not completely formatted before opening

## [1.0.0] - 2021-01-07

### Added

- Capability to edit OOXML document contents
- Get diff when ooxml document is modified from the outside

### Changed

- Rename menu option from "View Contents" to "Open OOXML Package"
- Save/cache OOXML parts to the file system

## [0.0.1] - 2019-08-27

### Added

- Support for viewing the contents of OOXML documents

[unreleased]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v2.0.0...HEAD
[2.0.0]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.3.1...v2.0.0
[1.3.1]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.3.0...v1.3.1
[1.3.0]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.2.2...v1.3.0
[1.2.2]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.1.1...v1.2.0
[1.1.1]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.0.2...v1.1.0
[1.0.2]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/yuenm18/ooxml-viewer-vscode/compare/v0.0.1...v1.0.0
[0.0.1]: https://github.com/yuenm18/ooxml-viewer-vscode/releases/tag/v0.0.1
