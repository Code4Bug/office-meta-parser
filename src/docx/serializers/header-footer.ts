import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { Header, Footer } from '../types.js';
import { serializeParagraph } from './paragraph.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function serializeHeader(header: Header): string {
  const children: ParsedNode[] = [];

  for (const para of header.content) {
    children.push(serializeParagraph(para));
  }

  const root: ParsedNode = {
    tag: 'w:hdr',
    attrs: {
      'xmlns:w': W_NS,
      'xmlns:r': R_NS,
    },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

export function serializeFooter(footer: Footer): string {
  const children: ParsedNode[] = [];

  for (const para of footer.content) {
    children.push(serializeParagraph(para));
  }

  const root: ParsedNode = {
    tag: 'w:ftr',
    attrs: {
      'xmlns:w': W_NS,
      'xmlns:r': R_NS,
    },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
