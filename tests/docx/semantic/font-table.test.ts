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

describe('DOCX semantic - fontTable', () => {
  it('extracts fonts from fontTable.xml', () => {
    const fontTableXml: ParsedNode = {
      tag: 'w:fonts',
      attrs: { 'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' },
      children: [
        {
          tag: 'w:font',
          attrs: { 'w:name': 'Times New Roman' },
          children: [
            { tag: 'w:charset', attrs: { 'w:val': '00' }, children: [] },
            { tag: 'w:family', attrs: { 'w:val': 'roman' }, children: [] },
            { tag: 'w:pitch', attrs: { 'w:val': 'variable' }, children: [] },
          ],
        },
        {
          tag: 'w:font',
          attrs: { 'w:name': '宋体' },
          children: [
            { tag: 'w:altName', attrs: { 'w:val': '汉仪书宋二KW' }, children: [] },
            { tag: 'w:charset', attrs: { 'w:val': '86' }, children: [] },
            { tag: 'w:family', attrs: { 'w:val': 'auto' }, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/fontTable.xml': fontTableXml });
    const semantic = rawToSemantic(raw);

    expect(semantic.fonts).toBeDefined();
    expect(semantic.fonts!.length).toBe(2);
    expect(semantic.fonts![0].name).toBe('Times New Roman');
    expect(semantic.fonts![0].charset).toBe('00');
    expect(semantic.fonts![0].family).toBe('roman');
    expect(semantic.fonts![0].pitch).toBe('variable');
    expect(semantic.fonts![1].name).toBe('宋体');
    expect(semantic.fonts![1].altName).toBe('汉仪书宋二KW');
    expect(semantic.fonts![1].charset).toBe('86');
  });

  it('returns undefined when no fontTable present', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.fonts).toBeUndefined();
  });
});
