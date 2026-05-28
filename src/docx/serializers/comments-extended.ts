import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { CommentExtended } from '../types.js';

const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

export function serializeCommentExts(commentExts: CommentExtended[]): string {
  const children: ParsedNode[] = commentExts.map(ce => ({
    tag: 'w15:commentEx',
    attrs: {
      'w15:paraId': ce.paraId,
      ...(ce.done ? { 'w15:done': '1' } : {}),
    },
    children: [],
  }));

  const root: ParsedNode = {
    tag: 'w15:commentsEx',
    attrs: { 'xmlns:w15': W15_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
