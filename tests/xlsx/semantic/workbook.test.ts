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

describe('XLSX semantic - workbook', () => {
  it('parses workbook with sheets', () => {
    const workbookXml: ParsedNode = {
      tag: 'workbook',
      attrs: {},
      children: [
        {
          tag: 'sheets',
          attrs: {},
          children: [
            { tag: 'sheet', attrs: { name: 'Sheet1', sheetId: '1', 'r:id': 'rId1' }, children: [] },
            { tag: 'sheet', attrs: { name: 'Sheet2', sheetId: '2', 'r:id': 'rId2' }, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/workbook.xml': workbookXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets).toHaveLength(2);
    expect(semantic.sheets[0].name).toBe('Sheet1');
    expect(semantic.sheets[1].name).toBe('Sheet2');
  });

  it('parses shared strings', () => {
    const sstXml: ParsedNode = {
      tag: 'sst',
      attrs: {},
      children: [
        { tag: 'si', attrs: {}, children: [{ tag: 't', attrs: {}, children: ['Hello'] }] },
        { tag: 'si', attrs: {}, children: [{ tag: 't', attrs: {}, children: ['World'] }] },
      ],
    };

    const raw = makeRaw({ 'xl/sharedStrings.xml': sstXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sharedStrings).toEqual([{ text: 'Hello' }, { text: 'World' }]);
  });
});
