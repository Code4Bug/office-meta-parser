export interface ZipEntry {
  path: string;
  data: ArrayBuffer;
}

export interface ParsedNode {
  tag: string;
  attrs: Record<string, string>;
  children: (ParsedNode | string)[];
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export interface ContentType {
  partName: string;
  contentType: string;
}

export interface RawDocument {
  entries: ZipEntry[];
  rels: Map<string, Relationship[]>;
  contentTypes: ContentType[];
  parts: Map<string, ParsedNode>;
}

export interface ParseResult<TSemantic> {
  raw: RawDocument;
  semantic: TSemantic;
}
