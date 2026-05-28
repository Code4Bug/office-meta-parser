import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { Comment } from '../types.js';
import { serializeParagraph } from './paragraph.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function serializeComments(comments: Comment[]): string {
  const children: ParsedNode[] = [];

  for (const comment of comments) {
    const attrs: Record<string, string> = {
      'w:id': comment.id,
      'w:author': comment.author,
      'w:date': comment.date,
    };
    if (comment.initials) {
      attrs['w:initials'] = comment.initials;
    }

    const commentChildren: ParsedNode[] = [];
    for (const para of comment.content) {
      commentChildren.push(serializeParagraph(para));
    }

    children.push({
      tag: 'w:comment',
      attrs,
      children: commentChildren,
    });
  }

  const root: ParsedNode = {
    tag: 'w:comments',
    attrs: {
      'xmlns:w': W_NS,
      'xmlns:r': R_NS,
    },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
