import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { Person } from '../types.js';

const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

export function serializePeople(people: Person[]): string {
  const children: ParsedNode[] = people.map(p => {
    const personChildren: ParsedNode[] = [];
    if (p.userId || p.providerId) {
      const attrs: Record<string, string> = {};
      if (p.providerId) attrs['w15:providerId'] = p.providerId;
      if (p.userId) attrs['w15:userId'] = p.userId;
      personChildren.push({ tag: 'w15:presenceInfo', attrs, children: [] });
    }
    return {
      tag: 'w15:person',
      attrs: { 'w15:author': p.author },
      children: personChildren,
    };
  });

  const root: ParsedNode = {
    tag: 'w15:people',
    attrs: { 'xmlns:w15': W15_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
