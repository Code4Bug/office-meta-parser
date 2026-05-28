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

describe('DOCX semantic - meta', () => {
  it('extracts meta from core.xml', () => {
    const coreXml: ParsedNode = {
      tag: 'cp:coreProperties',
      attrs: {
        'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
      },
      children: [
        { tag: 'dc:title', attrs: {}, children: ['My Document'] },
        { tag: 'dc:creator', attrs: {}, children: ['John Doe'] },
      ],
    };

    const raw = makeRaw({ 'docProps/core.xml': coreXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.meta.title).toBe('My Document');
    expect(semantic.meta.creator).toBe('John Doe');
  });
});
