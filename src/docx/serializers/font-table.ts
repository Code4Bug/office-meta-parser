import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { FontEntry } from '../types.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function serializeFontTable(fonts: FontEntry[]): string {
  const children: ParsedNode[] = fonts.map(f => {
    const fontChildren: ParsedNode[] = [];
    if (f.altName) fontChildren.push({ tag: 'w:altName', attrs: { 'w:val': f.altName }, children: [] });
    if (f.charset) fontChildren.push({ tag: 'w:charset', attrs: { 'w:val': f.charset }, children: [] });
    if (f.family) fontChildren.push({ tag: 'w:family', attrs: { 'w:val': f.family }, children: [] });
    if (f.pitch) fontChildren.push({ tag: 'w:pitch', attrs: { 'w:val': f.pitch }, children: [] });
    return {
      tag: 'w:font',
      attrs: { 'w:name': f.name },
      children: fontChildren,
    };
  });

  const root: ParsedNode = {
    tag: 'w:fonts',
    attrs: { 'xmlns:w': W_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
