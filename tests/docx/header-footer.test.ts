import { describe, it, expect } from 'vitest';
import { extractHeaders, extractFooters } from '../../src/docx/parsers/header-footer.js';
import type { RawDocument, ParsedNode, Relationship } from '../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>, rels?: Map<string, Relationship[]>): RawDocument {
  return {
    entries: [],
    rels: rels || new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('extractHeaders', () => {
  it('returns undefined when no headers', () => {
    const raw = makeRaw({});
    expect(extractHeaders(raw)).toBeUndefined();
  });

  it('parses header content', () => {
    const headerXml: ParsedNode = {
      tag: 'w:hdr',
      attrs: {},
      children: [
        {
          tag: 'w:p',
          attrs: {},
          children: [
            {
              tag: 'w:r',
              attrs: {},
              children: [
                { tag: 'w:t', attrs: {}, children: ['Page Header'] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/header1.xml': headerXml });
    const headers = extractHeaders(raw);
    expect(headers).toHaveLength(1);
    expect(headers![0].id).toBe('1');
    expect(headers![0].content).toHaveLength(1);
    expect(headers![0].content[0].runs[0].text).toBe('Page Header');
  });

  it('parses multiple headers', () => {
    const header1: ParsedNode = {
      tag: 'w:hdr',
      attrs: {},
      children: [
        {
          tag: 'w:p',
          attrs: {},
          children: [
            { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Default Header'] }] },
          ],
        },
      ],
    };

    const header2: ParsedNode = {
      tag: 'w:hdr',
      attrs: {},
      children: [
        {
          tag: 'w:p',
          attrs: {},
          children: [
            { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['First Page Header'] }] },
          ],
        },
      ],
    };

    const raw = makeRaw({
      'word/header1.xml': header1,
      'word/header2.xml': header2,
    });

    const headers = extractHeaders(raw);
    expect(headers).toHaveLength(2);
    expect(headers![0].content[0].runs[0].text).toBe('Default Header');
    expect(headers![1].content[0].runs[0].text).toBe('First Page Header');
  });
});

describe('extractFooters', () => {
  it('returns undefined when no footers', () => {
    const raw = makeRaw({});
    expect(extractFooters(raw)).toBeUndefined();
  });

  it('parses footer content', () => {
    const footerXml: ParsedNode = {
      tag: 'w:ftr',
      attrs: {},
      children: [
        {
          tag: 'w:p',
          attrs: {},
          children: [
            {
              tag: 'w:r',
              attrs: {},
              children: [
                { tag: 'w:t', attrs: {}, children: ['Page Footer'] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/footer1.xml': footerXml });
    const footers = extractFooters(raw);
    expect(footers).toHaveLength(1);
    expect(footers![0].id).toBe('1');
    expect(footers![0].content[0].runs[0].text).toBe('Page Footer');
  });

  it('parses footer with multiple paragraphs', () => {
    const footerXml: ParsedNode = {
      tag: 'w:ftr',
      attrs: {},
      children: [
        {
          tag: 'w:p',
          attrs: {},
          children: [
            { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Line 1'] }] },
          ],
        },
        {
          tag: 'w:p',
          attrs: {},
          children: [
            { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Line 2'] }] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/footer1.xml': footerXml });
    const footers = extractFooters(raw);
    expect(footers).toHaveLength(1);
    expect(footers![0].content).toHaveLength(2);
    expect(footers![0].content[0].runs[0].text).toBe('Line 1');
    expect(footers![0].content[1].runs[0].text).toBe('Line 2');
  });
});
