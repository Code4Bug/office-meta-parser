import { describe, it, expect } from 'vitest';
import { extractNumbering } from '../../src/docx/parsers/numbering.js';
import { serializeNumbering } from '../../src/docx/serializers/numbering.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

function makeRaw(numberingXml: ParsedNode): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map([['word/numbering.xml', numberingXml]]),
  };
}

const sampleNumberingXml: ParsedNode = {
  tag: 'w:numbering',
  attrs: {},
  children: [
    {
      tag: 'w:abstractNum',
      attrs: { 'w:abstractNumId': '0' },
      children: [
        {
          tag: 'w:lvl',
          attrs: { 'w:ilvl': '0' },
          children: [
            { tag: 'w:start', attrs: { 'w:val': '1' }, children: [] },
            { tag: 'w:numFmt', attrs: { 'w:val': 'decimal' }, children: [] },
            { tag: 'w:lvlText', attrs: { 'w:val': '%1.' }, children: [] },
            { tag: 'w:lvlJc', attrs: { 'w:val': 'left' }, children: [] },
            {
              tag: 'w:pPr',
              attrs: {},
              children: [
                { tag: 'w:ind', attrs: { 'w:left': '720', 'w:hanging': '360' }, children: [] },
              ],
            },
          ],
        },
      ],
    },
    {
      tag: 'w:num',
      attrs: { 'w:numId': '1' },
      children: [
        { tag: 'w:abstractNumId', attrs: { 'w:val': '0' }, children: [] },
      ],
    },
  ],
};

describe('DOCX numbering', () => {
  it('parses abstractNum with levels', () => {
    const defs = extractNumbering(makeRaw(sampleNumberingXml));
    expect(defs.abstractNums).toHaveLength(1);
    expect(defs.abstractNums[0].id).toBe('0');
    expect(defs.abstractNums[0].levels).toHaveLength(1);
    const lvl = defs.abstractNums[0].levels[0];
    expect(lvl.level).toBe(0);
    expect(lvl.numFmt).toBe('decimal');
    expect(lvl.lvlText).toBe('%1.');
    expect(lvl.start).toBe(1);
  });

  it('parses num → abstractNumId mapping', () => {
    const defs = extractNumbering(makeRaw(sampleNumberingXml));
    expect(defs.nums).toHaveLength(1);
    expect(defs.nums[0].id).toBe('1');
    expect(defs.nums[0].abstractNumId).toBe('0');
  });

  it('returns empty definitions when no numbering.xml', () => {
    const raw: RawDocument = {
      entries: [],
      rels: new Map(),
      contentTypes: [],
      parts: new Map(),
    };
    const defs = extractNumbering(raw);
    expect(defs.abstractNums).toHaveLength(0);
    expect(defs.nums).toHaveLength(0);
  });

  it('serializes numbering definitions round-trip', () => {
    const defs = extractNumbering(makeRaw(sampleNumberingXml));
    const xml = serializeNumbering(defs);
    expect(xml).toContain('w:numbering');
    expect(xml).toContain('w:abstractNum');
    expect(xml).toContain('w:abstractNumId="0"');
    expect(xml).toContain('w:numFmt');
    expect(xml).toContain('w:val="decimal"');
    expect(xml).toContain('w:num');
    expect(xml).toContain('w:numId="1"');
  });
});
