import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { Footnote } from '../types.js';
import { parseParagraph } from './paragraph.js';

export function extractEndnotes(raw: RawDocument): Footnote[] | undefined {
  const endnotesXml = raw.parts.get('word/endnotes.xml');
  if (!endnotesXml) return undefined;

  const endnotes: Footnote[] = [];

  for (const child of endnotesXml.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'w:endnote') continue;

    const type = child.attrs['w:type'];
    if (type === 'separator' || type === 'continuationSeparator') continue;

    const id = child.attrs['w:id'];
    if (!id) continue;

    const content = [];
    for (const fc of child.children) {
      if (typeof fc === 'string') continue;
      if (fc.tag === 'w:p') {
        content.push(parseParagraph(fc));
      }
    }

    const endnote: Footnote = { id, content };
    if (type) endnote.type = type;
    endnotes.push(endnote);
  }

  return endnotes.length > 0 ? endnotes : undefined;
}
