import { describe, it, expect } from 'vitest';
import { extractSharedStrings, extractSharedStringRichText } from '../../../src/xlsx/parsers/strings.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(sstXml: ParsedNode): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map([['xl/sharedStrings.xml', sstXml]]),
  };
}

describe('XLSX P0 semantic - shared strings rich text', () => {
  it('extracts plain text from shared strings', () => {
    const sst: ParsedNode = {
      tag: 'sst', attrs: {}, children: [
        { tag: 'si', attrs: {}, children: [{ tag: 't', attrs: {}, children: ['Hello'] }] },
        { tag: 'si', attrs: {}, children: [{ tag: 't', attrs: {}, children: ['World'] }] },
      ],
    };

    const raw = makeRaw(sst);
    const strings = extractSharedStrings(raw);
    expect(strings).toEqual([{ text: 'Hello' }, { text: 'World' }]);
  });

  it('extracts rich text from shared strings with formatting', () => {
    const sst: ParsedNode = {
      tag: 'sst', attrs: {}, children: [{
        tag: 'si', attrs: {}, children: [
          { tag: 'r', attrs: {}, children: [
            { tag: 'rPr', attrs: {}, children: [
              { tag: 'b', attrs: {}, children: [] },
              { tag: 'color', attrs: { rgb: 'FF0000' }, children: [] },
            ] },
            { tag: 't', attrs: {}, children: ['Bold Red'] },
          ] },
          { tag: 'r', attrs: {}, children: [
            { tag: 't', attrs: {}, children: [' Normal'] },
          ] },
        ],
      }],
    };

    const raw = makeRaw(sst);
    const richText = extractSharedStringRichText(raw, 0);
    expect(richText).toBeDefined();
    expect(richText).toHaveLength(2);
    expect(richText![0]).toEqual({ text: 'Bold Red', bold: true, color: 'FF0000' });
    expect(richText![1]).toEqual({ text: ' Normal' });
  });

  it('returns undefined for plain text shared string', () => {
    const sst: ParsedNode = {
      tag: 'sst', attrs: {}, children: [
        { tag: 'si', attrs: {}, children: [{ tag: 't', attrs: {}, children: ['Plain'] }] },
      ],
    };

    const raw = makeRaw(sst);
    const richText = extractSharedStringRichText(raw, 0);
    expect(richText).toBeUndefined();
  });

  it('handles rich text with multiple formatting properties', () => {
    const sst: ParsedNode = {
      tag: 'sst', attrs: {}, children: [{
        tag: 'si', attrs: {}, children: [{
          tag: 'r', attrs: {}, children: [
            { tag: 'rPr', attrs: {}, children: [
              { tag: 'b', attrs: {}, children: [] },
              { tag: 'i', attrs: {}, children: [] },
              { tag: 'u', attrs: {}, children: [] },
              { tag: 'strike', attrs: {}, children: [] },
              { tag: 'sz', attrs: { val: '14' }, children: [] },
              { tag: 'rFont', attrs: { val: 'Arial' }, children: [] },
            ] },
            { tag: 't', attrs: {}, children: ['Formatted'] },
          ],
        }],
      }],
    };

    const raw = makeRaw(sst);
    const richText = extractSharedStringRichText(raw, 0);
    expect(richText![0]).toEqual({
      text: 'Formatted',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      size: 14,
      font: 'Arial',
    });
  });

  it('concatenates plain r elements without rPr into plain text', () => {
    const sst: ParsedNode = {
      tag: 'sst', attrs: {}, children: [{
        tag: 'si', attrs: {}, children: [
          { tag: 'r', attrs: {}, children: [{ tag: 't', attrs: {}, children: ['Part1'] }] },
          { tag: 'r', attrs: {}, children: [{ tag: 't', attrs: {}, children: ['Part2'] }] },
        ],
      }],
    };

    const raw = makeRaw(sst);
    const strings = extractSharedStrings(raw);
    expect(strings[0].text).toBe('Part1Part2');
  });
});
