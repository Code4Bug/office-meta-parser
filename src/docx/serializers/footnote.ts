import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { Footnote } from '../types.js';
import { serializeParagraph } from './paragraph.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function serializeFootnotes(footnotes: Footnote[]): string {
  const children: ParsedNode[] = footnotes.map(fn => ({
    tag: 'w:footnote',
    attrs: {
      'w:id': fn.id,
      ...(fn.type ? { 'w:type': fn.type } : {}),
    },
    children: fn.content.map(serializeParagraph),
  }));

  const root: ParsedNode = {
    tag: 'w:footnotes',
    attrs: { 'xmlns:w': W_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
