import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { Comment } from '../types.js';
import { parseParagraph } from './paragraph.js';
import { findChild, getTextContent } from './utils.js';

export function extractComments(raw: RawDocument): Comment[] | undefined {
  const commentsXml = raw.parts.get('word/comments.xml');
  if (!commentsXml) return undefined;

  const comments: Comment[] = [];

  for (const child of commentsXml.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'w:comment') continue;

    const id = child.attrs['w:id'] || '';
    const author = child.attrs['w:author'] || '';
    const date = child.attrs['w:date'] || '';
    const initials = child.attrs['w:initials'] || '';

    const content = [];
    for (const para of child.children) {
      if (typeof para === 'string') continue;
      if (para.tag === 'w:p') {
        content.push(parseParagraph(para));
      }
    }

    const comment: Comment = {
      id,
      author,
      date,
      content,
    };
    if (initials) comment.initials = initials;

    comments.push(comment);
  }

  return comments.length > 0 ? comments : undefined;
}
