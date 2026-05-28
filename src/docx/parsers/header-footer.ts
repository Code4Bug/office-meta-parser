import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { Header, Footer, Paragraph } from '../types.js';
import { parseParagraph } from './paragraph.js';
import { findChild } from './utils.js';

export function extractHeaders(raw: RawDocument): Header[] | undefined {
  const headers: Header[] = [];

  for (const [path, xml] of raw.parts) {
    if (path.startsWith('word/header') && path.endsWith('.xml')) {
      const { id, type } = getHeaderFooterInfo(raw, path, 'header');
      const content = extractContent(xml);
      headers.push({ id, type, content });
    }
  }

  return headers.length > 0 ? headers : undefined;
}

export function extractFooters(raw: RawDocument): Footer[] | undefined {
  const footers: Footer[] = [];

  for (const [path, xml] of raw.parts) {
    if (path.startsWith('word/footer') && path.endsWith('.xml')) {
      const { id, type } = getHeaderFooterInfo(raw, path, 'footer');
      const content = extractContent(xml);
      footers.push({ id, type, content });
    }
  }

  return footers.length > 0 ? footers : undefined;
}

function getHeaderFooterInfo(raw: RawDocument, partPath: string, kind: 'header' | 'footer'): { id: string; type?: string } {
  // Look for the relationship in document.xml.rels
  const docRels = raw.rels.get('word/_rels/document.xml.rels');
  if (!docRels) {
    // Fallback to file number
    const num = partPath.match(new RegExp(`${kind}(\\d+)\\.xml`))?.[1] || '';
    return { id: num };
  }

  for (const rel of docRels) {
    const relTarget = 'word/' + rel.target;
    if (relTarget === partPath) {
      // Extract type from relationship type URI
      let type: string | undefined;
      if (rel.type.includes('first')) type = 'first';
      else if (rel.type.includes('even')) type = 'even';
      else type = 'default';

      return { id: rel.id, type };
    }
  }

  // Fallback to file number
  const num = partPath.match(new RegExp(`${kind}(\\d+)\\.xml`))?.[1] || '';
  return { id: num };
}

function extractContent(node: ParsedNode): Paragraph[] {
  const body = findChild(node, 'w:hdr') || findChild(node, 'w:ftr') || node;
  const paragraphs: Paragraph[] = [];

  function collectParagraphs(n: ParsedNode) {
    for (const child of n.children) {
      if (typeof child === 'string') continue;
      if (child.tag === 'w:p') {
        paragraphs.push(parseParagraph(child));
      } else {
        // 递归查找嵌套在文本框等结构中的段落
        collectParagraphs(child);
      }
    }
  }

  collectParagraphs(body);
  return paragraphs;
}
