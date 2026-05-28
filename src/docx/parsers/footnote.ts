import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { Footnote, Paragraph } from '../types.js';
import { parseParagraph } from './paragraph.js';

export function extractFootnotes(raw: RawDocument): Footnote[] {
  const footnotesXml = raw.parts.get('word/footnotes.xml');
  if (!footnotesXml) return [];

  const footnotes: Footnote[] = [];

  for (const child of footnotesXml.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:footnote') {
      const type = child.attrs['w:type'];
      // 跳过内置的分隔符和延续符
      if (type === 'separator' || type === 'continuationSeparator') continue;

      const id = child.attrs['w:id'];
      if (!id) continue;

      const content: Paragraph[] = [];
      for (const fc of child.children) {
        if (typeof fc === 'string') continue;
        if (fc.tag === 'w:p') {
          content.push(parseParagraph(fc));
        }
      }

      const footnote: Footnote = { id, content };
      if (type) footnote.type = type;
      footnotes.push(footnote);
    }
  }

  return footnotes;
}
