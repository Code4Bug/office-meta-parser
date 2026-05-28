import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { CommentExtended } from '../types.js';

export function extractCommentExts(raw: RawDocument): CommentExtended[] | undefined {
  const xml = raw.parts.get('word/commentsExtended.xml');
  if (!xml) return undefined;

  const result: CommentExtended[] = [];

  for (const child of xml.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'w15:commentEx') continue;

    const paraId = child.attrs['w15:paraId'];
    if (!paraId) continue;

    const ext: CommentExtended = { paraId };
    if (child.attrs['w15:done'] === '1' || child.attrs['w15:done'] === 'true') {
      ext.done = true;
    }

    result.push(ext);
  }

  return result.length > 0 ? result : undefined;
}
