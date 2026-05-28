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

describe('DOCX semantic - people', () => {
  it('extracts people from people.xml', () => {
    const xml: ParsedNode = {
      tag: 'w15:people',
      attrs: { 'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml' },
      children: [
        {
          tag: 'w15:person',
          attrs: { 'w15:author': '张三' },
          children: [{
            tag: 'w15:presenceInfo',
            attrs: { 'w15:providerId': 'WPS Office', 'w15:userId': '12345' },
            children: [],
          }],
        },
        {
          tag: 'w15:person',
          attrs: { 'w15:author': '李四' },
          children: [],
        },
      ],
    };

    const raw = makeRaw({ 'word/people.xml': xml });
    const semantic = rawToSemantic(raw);

    expect(semantic.people).toBeDefined();
    expect(semantic.people!.length).toBe(2);
    expect(semantic.people![0].author).toBe('张三');
    expect(semantic.people![0].userId).toBe('12345');
    expect(semantic.people![0].providerId).toBe('WPS Office');
    expect(semantic.people![1].author).toBe('李四');
    expect(semantic.people![1].userId).toBeUndefined();
  });

  it('returns undefined when no people present', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.people).toBeUndefined();
  });
});
