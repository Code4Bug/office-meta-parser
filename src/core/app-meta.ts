import { serializeXml } from './xml.js';
import type { ParsedNode } from './types.js';

export interface AppMeta {
  template?: string;
  totalTime?: number;
  pages?: number;
  words?: number;
  characters?: number;
  charactersWithSpaces?: number;
  application?: string;
  docSecurity?: number;
  scaleCrop?: boolean;
  company?: string;
  linksUpToDate?: boolean;
  sharedDoc?: boolean;
  hyperlinksChanged?: boolean;
  appVersion?: string;
}

const EP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';

function getTextContent(node: ParsedNode): string {
  for (const child of node.children) {
    if (typeof child === 'string') return child;
  }
  return '';
}

export function parseAppMeta(node: ParsedNode): AppMeta {
  const meta: AppMeta = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    const text = getTextContent(child);
    switch (child.tag) {
      case 'Template': meta.template = text; break;
      case 'TotalTime': meta.totalTime = parseInt(text, 10); break;
      case 'Pages': meta.pages = parseInt(text, 10); break;
      case 'Words': meta.words = parseInt(text, 10); break;
      case 'Characters': meta.characters = parseInt(text, 10); break;
      case 'CharactersWithSpaces': meta.charactersWithSpaces = parseInt(text, 10); break;
      case 'Application': meta.application = text; break;
      case 'DocSecurity': meta.docSecurity = parseInt(text, 10); break;
      case 'ScaleCrop': meta.scaleCrop = text === 'true'; break;
      case 'Company': meta.company = text; break;
      case 'LinksUpToDate': meta.linksUpToDate = text === 'true'; break;
      case 'SharedDoc': meta.sharedDoc = text === 'true'; break;
      case 'HyperlinksChanged': meta.hyperlinksChanged = text === 'true'; break;
      case 'AppVersion': meta.appVersion = text; break;
    }
  }

  return meta;
}

export function serializeAppMeta(meta: AppMeta): string {
  const children: ParsedNode[] = [];

  if (meta.template) {
    children.push({ tag: 'Template', attrs: {}, children: [meta.template] });
  }
  if (meta.totalTime !== undefined) {
    children.push({ tag: 'TotalTime', attrs: {}, children: [String(meta.totalTime)] });
  }
  if (meta.pages !== undefined) {
    children.push({ tag: 'Pages', attrs: {}, children: [String(meta.pages)] });
  }
  if (meta.words !== undefined) {
    children.push({ tag: 'Words', attrs: {}, children: [String(meta.words)] });
  }
  if (meta.characters !== undefined) {
    children.push({ tag: 'Characters', attrs: {}, children: [String(meta.characters)] });
  }
  if (meta.charactersWithSpaces !== undefined) {
    children.push({ tag: 'CharactersWithSpaces', attrs: {}, children: [String(meta.charactersWithSpaces)] });
  }
  if (meta.application) {
    children.push({ tag: 'Application', attrs: {}, children: [meta.application] });
  }
  if (meta.docSecurity !== undefined) {
    children.push({ tag: 'DocSecurity', attrs: {}, children: [String(meta.docSecurity)] });
  }
  if (meta.scaleCrop !== undefined) {
    children.push({ tag: 'ScaleCrop', attrs: {}, children: [String(meta.scaleCrop)] });
  }
  if (meta.company) {
    children.push({ tag: 'Company', attrs: {}, children: [meta.company] });
  }
  if (meta.linksUpToDate !== undefined) {
    children.push({ tag: 'LinksUpToDate', attrs: {}, children: [String(meta.linksUpToDate)] });
  }
  if (meta.sharedDoc !== undefined) {
    children.push({ tag: 'SharedDoc', attrs: {}, children: [String(meta.sharedDoc)] });
  }
  if (meta.hyperlinksChanged !== undefined) {
    children.push({ tag: 'HyperlinksChanged', attrs: {}, children: [String(meta.hyperlinksChanged)] });
  }
  if (meta.appVersion) {
    children.push({ tag: 'AppVersion', attrs: {}, children: [meta.appVersion] });
  }

  const root: ParsedNode = {
    tag: 'Properties',
    attrs: { xmlns: EP_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
