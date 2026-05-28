import { parseXml, serializeXml } from './xml.js';
import type { Relationship, ParsedNode } from './types.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export function parseRels(xml: string): Relationship[] {
  const root = parseXml(xml);
  if (root.tag !== 'Relationships') {
    throw new Error(`Expected Relationships root, got ${root.tag}`);
  }

  return root.children
    .filter((c): c is ParsedNode => typeof c !== 'string' && c.tag === 'Relationship')
    .map(node => ({
      id: node.attrs['Id'] || '',
      type: node.attrs['Type'] || '',
      target: node.attrs['Target'] || '',
      targetMode: node.attrs['TargetMode'],
    }));
}

export function serializeRels(rels: Relationship[]): string {
  const root: ParsedNode = {
    tag: 'Relationships',
    attrs: { xmlns: RELS_NS },
    children: rels.map(rel => {
      const attrs: Record<string, string> = {
        Id: rel.id,
        Type: rel.type,
        Target: rel.target,
      };
      if (rel.targetMode) {
        attrs.TargetMode = rel.targetMode;
      }
      return {
        tag: 'Relationship',
        attrs,
        children: [],
      };
    }),
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
