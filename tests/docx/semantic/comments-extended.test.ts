import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/docx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('DOCX semantic - commentsExtended', () => {
  it('extracts comment extensions', () => {
    const xml: ParsedNode = {
      tag: 'w15:commentsEx',
      attrs: { 'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml' },
      children: [
        { tag: 'w15:commentEx', attrs: { 'w15:paraId': '6D3DE01F', 'w15:done': '0' }, children: [] },
        { tag: 'w15:commentEx', attrs: { 'w15:paraId': '7E744B72', 'w15:done': '1' }, children: [] },
      ],
    };

    const raw = makeRaw({ 'word/commentsExtended.xml': xml });
    const semantic = rawToSemantic(raw);

    expect(semantic.commentExts).toBeDefined();
    expect(semantic.commentExts!.length).toBe(2);
    expect(semantic.commentExts![0].paraId).toBe('6D3DE01F');
    expect(semantic.commentExts![0].done).toBeUndefined();
    expect(semantic.commentExts![1].paraId).toBe('7E744B72');
    expect(semantic.commentExts![1].done).toBe(true);
  });

  it('returns undefined when no commentsExtended present', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.commentExts).toBeUndefined();
  });
});
