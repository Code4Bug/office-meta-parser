import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { Footnote } from '../types.js';
import { serializeParagraph } from './paragraph.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function serializeEndnotes(endnotes: Footnote[]): string {
  const children: ParsedNode[] = endnotes.map(en => ({
    tag: 'w:endnote',
    attrs: {
      'w:id': en.id,
      ...(en.type ? { 'w:type': en.type } : {}),
    },
    children: en.content.map(serializeParagraph),
  }));

  const root: ParsedNode = {
    tag: 'w:endnotes',
    attrs: { 'xmlns:w': W_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
