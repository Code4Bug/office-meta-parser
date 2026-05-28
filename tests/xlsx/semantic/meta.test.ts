import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/xlsx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('XLSX semantic - meta', () => {
  it('extracts meta from core.xml', () => {
    const coreXml: ParsedNode = {
      tag: 'cp:coreProperties',
      attrs: {},
      children: [
        { tag: 'dc:title', attrs: {}, children: ['My Spreadsheet'] },
        { tag: 'dc:creator', attrs: {}, children: ['Jane'] },
      ],
    };

    const raw = makeRaw({ 'docProps/core.xml': coreXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.meta.title).toBe('My Spreadsheet');
    expect(semantic.meta.creator).toBe('Jane');
  });

  it('handles empty workbook', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets).toEqual([]);
    expect(semantic.sharedStrings).toEqual([]);
  });
});
