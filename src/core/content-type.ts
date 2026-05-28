import { parseXml, serializeXml } from './xml.js';
import type { ContentType, ParsedNode } from './types.js';

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

export function parseContentTypes(xml: string): ContentType[] {
  const root = parseXml(xml);
  if (root.tag !== 'Types') {
    throw new Error(`Expected Types root, got ${root.tag}`);
  }

  const result: ContentType[] = [];

  for (const child of root.children) {
    if (typeof child === 'string') continue;

    if (child.tag === 'Default') {
      result.push({
        partName: '.' + (child.attrs['Extension'] || ''),
        contentType: child.attrs['ContentType'] || '',
      });
    } else if (child.tag === 'Override') {
      result.push({
        partName: child.attrs['PartName'] || '',
        contentType: child.attrs['ContentType'] || '',
      });
    }
  }

  return result;
}

export function serializeContentTypes(types: ContentType[]): string {
  const children: ParsedNode[] = types.map(ct => {
    if (ct.partName.startsWith('.')) {
      const attrs: Record<string, string> = { Extension: ct.partName.slice(1), ContentType: ct.contentType };
      return { tag: 'Default', attrs, children: [] };
    }
    const attrs: Record<string, string> = { PartName: ct.partName, ContentType: ct.contentType };
    return { tag: 'Override', attrs, children: [] };
  });

  const root: ParsedNode = {
    tag: 'Types',
    attrs: { xmlns: CT_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
