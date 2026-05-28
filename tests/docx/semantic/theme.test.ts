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

describe('DOCX semantic - theme', () => {
  it('extracts color scheme and fonts from theme', () => {
    const themeXml: ParsedNode = {
      tag: 'a:theme',
      attrs: { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main', name: 'Office' },
      children: [{
        tag: 'a:themeElements',
        attrs: {},
        children: [
          {
            tag: 'a:clrScheme',
            attrs: { name: '' },
            children: [
              { tag: 'a:dk1', attrs: {}, children: [{ tag: 'a:srgbClr', attrs: { val: '000000' }, children: [] }] },
              { tag: 'a:lt1', attrs: {}, children: [{ tag: 'a:srgbClr', attrs: { val: 'FFFFFF' }, children: [] }] },
              { tag: 'a:accent1', attrs: {}, children: [{ tag: 'a:srgbClr', attrs: { val: '4F81BD' }, children: [] }] },
            ],
          },
          {
            tag: 'a:fontScheme',
            attrs: { name: '' },
            children: [
              { tag: 'a:majorFont', attrs: {}, children: [{ tag: 'a:latin', attrs: { typeface: 'Arial' }, children: [] }] },
              { tag: 'a:minorFont', attrs: {}, children: [{ tag: 'a:latin', attrs: { typeface: 'Calibri' }, children: [] }] },
            ],
          },
        ],
      }],
    };

    const raw = makeRaw({ 'word/theme/theme1.xml': themeXml });
    const semantic = rawToSemantic(raw);

    expect(semantic.theme).toBeDefined();
    expect(semantic.theme!.colorScheme).toBeDefined();
    expect(semantic.theme!.colorScheme!['dk1']).toBe('000000');
    expect(semantic.theme!.colorScheme!['lt1']).toBe('FFFFFF');
    expect(semantic.theme!.colorScheme!['accent1']).toBe('4F81BD');
    expect(semantic.theme!.majorFont).toBe('Arial');
    expect(semantic.theme!.minorFont).toBe('Calibri');
  });

  it('returns undefined when no theme present', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.theme).toBeUndefined();
  });
});
