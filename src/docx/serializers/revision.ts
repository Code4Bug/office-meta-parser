import type { ParsedNode } from '../../core/types.js';
import type { Revision } from '../types.js';

export function serializeRevisionMark(revision: Revision): ParsedNode {
  const tag = revision.type === 'insert' ? 'w:ins' :
    revision.type === 'delete' ? 'w:del' : 'w:rPrChange';

  return {
    tag,
    attrs: {
      'w:author': revision.author,
      'w:date': revision.date,
    },
    children: [],
  };
}
