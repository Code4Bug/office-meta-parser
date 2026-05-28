import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/docx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(sectPrChildren: ParsedNode[]): RawDocument {
  const documentXml: ParsedNode = {
    tag: 'w:document',
    attrs: {},
    children: [{
      tag: 'w:body',
      attrs: {},
      children: [
        { tag: 'w:sectPr', attrs: {}, children: sectPrChildren },
      ],
    }],
  };
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map([['word/document.xml', documentXml]]),
  };
}

describe('DOCX parser - section properties P0', () => {
  it('parses landscape orientation', () => {
    const sem = rawToSemantic(makeRaw([
      { tag: 'w:pgSz', attrs: { 'w:orient': 'landscape' }, children: [] },
    ]));
    expect(sem.body.sectionProperties?.orientation).toBe('landscape');
  });

  it('parses page number format', () => {
    const sem = rawToSemantic(makeRaw([
      { tag: 'w:pgNumType', attrs: { 'w:fmt': 'lowerRoman', 'w:start': '3' }, children: [] },
    ]));
    expect(sem.body.sectionProperties?.pageNumberFormat).toBe('lowerRoman');
    expect(sem.body.sectionProperties?.pageNumberStart).toBe(3);
  });

  it('parses titlePage', () => {
    const sem = rawToSemantic(makeRaw([
      { tag: 'w:titlePg', attrs: {}, children: [] },
    ]));
    expect(sem.body.sectionProperties?.titlePage).toBe(true);
  });

  it('parses evenAndOddHeaders', () => {
    const sem = rawToSemantic(makeRaw([
      { tag: 'w:evenAndOddHeaders', attrs: {}, children: [] },
    ]));
    expect(sem.body.sectionProperties?.evenAndOddHeaders).toBe(true);
  });

  it('parses vertical alignment', () => {
    const sem = rawToSemantic(makeRaw([
      { tag: 'w:vAlign', attrs: { 'w:val': 'center' }, children: [] },
    ]));
    expect(sem.body.sectionProperties?.verticalAlign).toBe('center');
  });
});
