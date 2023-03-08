import { FileNode } from '../tree-view/ooxml-tree-view-provider';

export interface OOXMLCommand {
  command: string;
  fileNode: FileNode;
}

export class RemoveOOXMLCommand implements OOXMLCommand {
  command = 'ooxmlViewer.removeOoxmlPackage';

  constructor(public fileNode: FileNode) {}
}
