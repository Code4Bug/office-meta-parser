import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { Hyperlink, TextRun } from '../types.js';
import { parseRun } from './paragraph.js';

export function parseHyperlink(raw: RawDocument, node: ParsedNode): Hyperlink {
  const runs: TextRun[] = [];
  const relId = node.attrs['r:id'] || '';
  const tooltip = node.attrs['w:tooltip'] || '';

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:r') {
      runs.push(parseRun(child));
    }
  }

  const docRels = raw.rels.get('word/_rels/document.xml.rels');
  let url = '';
  if (docRels) {
    const rel = docRels.find(r => r.id === relId);
    if (rel) url = rel.target;
  }

  const result: Hyperlink = {
    type: 'hyperlink',
    relationshipId: relId,
    runs,
  };
  if (url) result.url = url;
  if (tooltip) result.tooltip = tooltip;
  return result;
}
