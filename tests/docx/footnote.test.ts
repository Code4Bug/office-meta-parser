import { describe, it, expect } from 'vitest';
import { extractFootnotes } from '../../src/docx/parsers/footnote.js';
import { serializeFootnotes } from '../../src/docx/serializers/footnote.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

function makeRaw(footnotesXml: ParsedNode): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map([['word/footnotes.xml', footnotesXml]]),
  };
}

const sampleFootnotesXml: ParsedNode = {
  tag: 'w:footnotes',
  attrs: {},
  children: [
    {
      tag: 'w:footnote',
      attrs: { 'w:id': '1', 'w:type': 'normal' },
      children: [
        {
          tag: 'w:p',
          attrs: {},
          children: [
            {
              tag: 'w:r',
              attrs: {},
              children: [{ tag: 'w:t', attrs: {}, children: ['Footnote text'] }],
            },
          ],
        },
      ],
    },
  ],
};

describe('DOCX footnotes', () => {
  it('parses footnotes with id and content', () => {
    const footnotes = extractFootnotes(makeRaw(sampleFootnotesXml));
    expect(footnotes).toHaveLength(1);
    expect(footnotes[0].id).toBe('1');
    expect(footnotes[0].content).toHaveLength(1);
    expect(footnotes[0].content[0].runs[0].text).toBe('Footnote text');
  });

  it('returns empty array when no footnotes.xml', () => {
    const raw: RawDocument = {
      entries: [],
      rels: new Map(),
      contentTypes: [],
      parts: new Map(),
    };
    expect(extractFootnotes(raw)).toHaveLength(0);
  });

  it('serializes footnotes', () => {
    const footnotes = extractFootnotes(makeRaw(sampleFootnotesXml));
    const xml = serializeFootnotes(footnotes);
    expect(xml).toContain('w:footnotes');
    expect(xml).toContain('w:footnote');
    expect(xml).toContain('w:id="1"');
    expect(xml).toContain('Footnote text');
  });
});
