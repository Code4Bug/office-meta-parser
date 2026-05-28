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

describe('DOCX semantic - endnotes', () => {
  it('extracts endnotes from endnotes.xml', () => {
    const xml: ParsedNode = {
      tag: 'w:endnotes',
      attrs: { 'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' },
      children: [
        // 内置分隔符应跳过
        {
          tag: 'w:endnote',
          attrs: { 'w:type': 'separator', 'w:id': '0' },
          children: [{
            tag: 'w:p',
            attrs: {},
            children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:separator', attrs: {}, children: [] }] }],
          }],
        },
        // 实际尾注
        {
          tag: 'w:endnote',
          attrs: { 'w:id': '2' },
          children: [{
            tag: 'w:p',
            attrs: {},
            children: [{
              tag: 'w:r',
              attrs: {},
              children: [{ tag: 'w:t', attrs: {}, children: ['这是尾注内容'] }],
            }],
          }],
        },
      ],
    };

    const raw = makeRaw({ 'word/endnotes.xml': xml });
    const semantic = rawToSemantic(raw);

    expect(semantic.endnotes).toBeDefined();
    expect(semantic.endnotes!.length).toBe(1);
    expect(semantic.endnotes![0].id).toBe('2');
    expect(semantic.endnotes![0].content.length).toBe(1);
    expect(semantic.endnotes![0].content[0].runs[0].text).toBe('这是尾注内容');
  });

  it('returns undefined when no endnotes present', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.endnotes).toBeUndefined();
  });
});
