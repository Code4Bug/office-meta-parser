import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/pptx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('PPTX semantic - meta', () => {
  it('extracts meta from core.xml', () => {
    const coreXml: ParsedNode = {
      tag: 'cp:coreProperties',
      attrs: {},
      children: [
        { tag: 'dc:title', attrs: {}, children: ['My Presentation'] },
        { tag: 'dc:creator', attrs: {}, children: ['John'] },
      ],
    };

    const raw = makeRaw({ 'docProps/core.xml': coreXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.meta.title).toBe('My Presentation');
    expect(semantic.meta.creator).toBe('John');
  });

  it('handles empty presentation', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.slides).toEqual([]);
    expect(semantic.masters).toEqual([]);
    expect(semantic.layouts).toEqual([]);
  });
});
