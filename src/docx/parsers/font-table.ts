import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { FontEntry } from '../types.js';

export function extractFonts(raw: RawDocument): FontEntry[] | undefined {
  const fontTableXml = raw.parts.get('word/fontTable.xml');
  if (!fontTableXml) return undefined;

  const fonts: FontEntry[] = [];

  for (const child of fontTableXml.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'w:font') continue;

    const name = child.attrs['w:name'];
    if (!name) continue;

    const entry: FontEntry = { name };

    for (const fc of child.children) {
      if (typeof fc === 'string') continue;
      switch (fc.tag) {
        case 'w:altName': entry.altName = fc.attrs['w:val']; break;
        case 'w:charset': entry.charset = fc.attrs['w:val']; break;
        case 'w:family': entry.family = fc.attrs['w:val']; break;
        case 'w:pitch': entry.pitch = fc.attrs['w:val']; break;
      }
    }

    fonts.push(entry);
  }

  return fonts.length > 0 ? fonts : undefined;
}
