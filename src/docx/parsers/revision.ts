import type { ParsedNode, RawDocument } from '../../core/types.js';
import type { Revision } from '../types.js';

export function extractRevisions(raw: RawDocument): Revision[] | undefined {
  const documentXml = raw.parts.get('word/document.xml');
  if (!documentXml) return undefined;

  const revisions: Revision[] = [];
  collectRevisions(documentXml, revisions);

  return revisions.length > 0 ? revisions : undefined;
}

function collectRevisions(node: ParsedNode | string, revisions: Revision[]): void {
  if (typeof node === 'string') return;

  if (node.tag === 'w:ins') {
    const rev = parseRevision(node, 'insert');
    if (rev) revisions.push(rev);
  } else if (node.tag === 'w:del') {
    const rev = parseRevision(node, 'delete');
    if (rev) revisions.push(rev);
  } else if (node.tag === 'w:rPrChange') {
    const rev = parseRevision(node, 'formatChange');
    if (rev) revisions.push(rev);
  } else if (node.tag === 'w:pPrChange') {
    const rev = parseRevision(node, 'formatChange');
    if (rev) revisions.push(rev);
  }

  for (const child of node.children) {
    collectRevisions(child, revisions);
  }
}

function parseRevision(node: ParsedNode, type: Revision['type']): Revision | undefined {
  const author = node.attrs['w:author'];
  const date = node.attrs['w:date'];
  if (!author || !date) return undefined;

  return { type, author, date };
}
