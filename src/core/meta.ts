import { serializeXml } from './xml.js';
import { toJSON, toJSONString, saveToJSON } from './io.js';
import type { ParsedNode } from './types.js';

export interface DocumentMeta {
  title?: string;
  subject?: string;
  creator?: string;
  description?: string;
  keywords?: string;
  lastModifiedBy?: string;
  created?: string;
  modified?: string;
  revision?: string;
  category?: string;
}

type MetaHolder = { meta: DocumentMeta };

function touch(obj: MetaHolder): void {
  obj.meta.modified = new Date().toISOString();
}

export function updateTitle<T extends MetaHolder>(doc: T, title: string): void {
  doc.meta.title = title;
  touch(doc);
}

export function updateSubject<T extends MetaHolder>(doc: T, subject: string): void {
  doc.meta.subject = subject;
  touch(doc);
}

export function updateCreator<T extends MetaHolder>(doc: T, creator: string): void {
  doc.meta.creator = creator;
  touch(doc);
}

export function updateDescription<T extends MetaHolder>(doc: T, description: string): void {
  doc.meta.description = description;
  touch(doc);
}

export function updateKeywords<T extends MetaHolder>(doc: T, keywords: string): void {
  doc.meta.keywords = keywords;
  touch(doc);
}

export function updateLastModifiedBy<T extends MetaHolder>(doc: T, lastModifiedBy: string): void {
  doc.meta.lastModifiedBy = lastModifiedBy;
  touch(doc);
}

export function updateCategory<T extends MetaHolder>(doc: T, category: string): void {
  doc.meta.category = category;
  touch(doc);
}

function getTextContent(node: ParsedNode): string {
  for (const child of node.children) {
    if (typeof child === 'string') return child;
  }
  return '';
}

export function parseMeta(node: ParsedNode): DocumentMeta {
  const meta: DocumentMeta = {};
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    const text = getTextContent(child);
    switch (child.tag) {
      case 'dc:title': meta.title = text; break;
      case 'dc:subject': meta.subject = text; break;
      case 'dc:creator': meta.creator = text; break;
      case 'dc:description': meta.description = text; break;
      case 'cp:keywords': meta.keywords = text; break;
      case 'cp:lastModifiedBy': meta.lastModifiedBy = text; break;
      case 'dcterms:created': meta.created = text; break;
      case 'dcterms:modified': meta.modified = text; break;
      case 'cp:revision': meta.revision = text; break;
      case 'cp:category': meta.category = text; break;
    }
  }
  return meta;
}

const CP_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC_NS = 'http://purl.org/dc/elements/1.1/';
const DCTERMS_NS = 'http://purl.org/dc/terms/';
const XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance';

export function serializeMeta(meta: DocumentMeta): string {
  const children: ParsedNode[] = [];

  if (meta.title) {
    children.push({ tag: 'dc:title', attrs: {}, children: [meta.title] });
  }
  if (meta.subject) {
    children.push({ tag: 'dc:subject', attrs: {}, children: [meta.subject] });
  }
  if (meta.creator) {
    children.push({ tag: 'dc:creator', attrs: {}, children: [meta.creator] });
  }
  if (meta.description) {
    children.push({ tag: 'dc:description', attrs: {}, children: [meta.description] });
  }
  if (meta.keywords) {
    children.push({ tag: 'cp:keywords', attrs: {}, children: [meta.keywords] });
  }
  if (meta.lastModifiedBy) {
    children.push({ tag: 'cp:lastModifiedBy', attrs: {}, children: [meta.lastModifiedBy] });
  }
  if (meta.created) {
    children.push({
      tag: 'dcterms:created',
      attrs: { 'xsi:type': 'dcterms:W3CDTF' },
      children: [meta.created],
    });
  }
  if (meta.modified) {
    children.push({
      tag: 'dcterms:modified',
      attrs: { 'xsi:type': 'dcterms:W3CDTF' },
      children: [meta.modified],
    });
  }
  if (meta.revision) {
    children.push({ tag: 'cp:revision', attrs: {}, children: [meta.revision] });
  }
  if (meta.category) {
    children.push({ tag: 'cp:category', attrs: {}, children: [meta.category] });
  }

  const root: ParsedNode = {
    tag: 'cp:coreProperties',
    attrs: {
      'xmlns:cp': CP_NS,
      'xmlns:dc': DC_NS,
      'xmlns:dcterms': DCTERMS_NS,
      'xmlns:xsi': XSI_NS,
    },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

/**
 * 为任意持有 meta 的类型生成全套元数据 + JSON 操作
 */
export function createMetaOps<T extends MetaHolder>() {
  return {
    updateTitle: (doc: T, v: string) => updateTitle(doc, v),
    updateSubject: (doc: T, v: string) => updateSubject(doc, v),
    updateCreator: (doc: T, v: string) => updateCreator(doc, v),
    updateDescription: (doc: T, v: string) => updateDescription(doc, v),
    updateKeywords: (doc: T, v: string) => updateKeywords(doc, v),
    updateLastModifiedBy: (doc: T, v: string) => updateLastModifiedBy(doc, v),
    updateCategory: (doc: T, v: string) => updateCategory(doc, v),
    toJSON: (doc: T): T => toJSON(doc),
    toJSONString: (doc: T, space?: number): string => toJSONString(doc, space),
    saveJSON: (doc: T, path: string, space?: number): Promise<void> => saveToJSON(doc, path, space),
  };
}
