import type { RawDocument } from '../../core/types.js';
import type { Hyperlink } from '../types.js';
import { findChild } from './utils.js';

export function extractHyperlinks(raw: RawDocument, sheetPath: string): Hyperlink[] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  const hyperlinksNode = findChild(worksheetXml, 'hyperlinks');
  if (!hyperlinksNode) return [];

  // Get sheet rels to resolve relationship IDs
  const sheetRelsPath = sheetPath.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
  const sheetRels = raw.rels.get(sheetRelsPath);

  const hyperlinks: Hyperlink[] = [];

  for (const child of hyperlinksNode.children) {
    if (typeof child === 'string' || child.tag !== 'hyperlink') continue;

    const ref = child.attrs['ref'];
    if (!ref) continue;

    let url = '';
    const relId = child.attrs['r:id'];
    const location = child.attrs['location'];

    if (location) {
      url = location;
    } else if (relId && sheetRels) {
      const rel = sheetRels.find(r => r.id === relId);
      if (rel) url = rel.target;
    }

    const tooltip = child.attrs['tooltip'];

    const hyperlink: Hyperlink = { ref, url };
    if (tooltip) hyperlink.tooltip = tooltip;

    hyperlinks.push(hyperlink);
  }

  return hyperlinks;
}
