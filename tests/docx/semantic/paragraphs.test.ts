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

describe('DOCX semantic - paragraphs', () => {
  it('parses a simple paragraph with text run', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
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
                    { tag: 'w:t', attrs: {}, children: ['Hello World'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.body.blocks).toHaveLength(1);
    const para = semantic.body.blocks[0] as any;
    expect(para.type).toBe('paragraph');
    expect(para.runs).toHaveLength(1);
    expect(para.runs[0].text).toBe('Hello World');
  });

  it('parses bold and italic runs', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
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
                    {
                      tag: 'w:rPr',
                      attrs: {},
                      children: [
                        { tag: 'w:b', attrs: {}, children: [] },
                        { tag: 'w:i', attrs: {}, children: [] },
                      ],
                    },
                    { tag: 'w:t', attrs: {}, children: ['Bold Italic'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.runs[0].bold).toBe(true);
    expect(para.runs[0].italic).toBe(true);
  });

  it('handles paragraph with no runs (empty paragraph)', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            { tag: 'w:p', attrs: {}, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.body.blocks).toHaveLength(1);
    expect((semantic.body.blocks[0] as any).runs).toEqual([]);
  });

  it('handles run with mixed content (text + tab + text)', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
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
                    { tag: 'w:t', attrs: {}, children: ['Hello'] },
                    { tag: 'w:tab', attrs: {}, children: [] },
                    { tag: 'w:t', attrs: {}, children: ['World'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.runs).toHaveLength(1);
    expect(para.runs[0].text).toBe('Hello');
  });

  it('parses paragraph alignment', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:pPr', attrs: {}, children: [{ tag: 'w:jc', attrs: { 'w:val': 'center' }, children: [] }] },
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Centered'] }] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.properties.alignment).toBe('center');
  });

  it('parses paragraph indent', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:pPr', attrs: {}, children: [{ tag: 'w:ind', attrs: { 'w:left': '720', 'w:firstLine': '480' }, children: [] }] },
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Indented'] }] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.properties.indent.left).toBe(720);
    expect(para.properties.indent.firstLine).toBe(480);
  });

  it('parses paragraph spacing', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:pPr', attrs: {}, children: [{ tag: 'w:spacing', attrs: { 'w:before': '120', 'w:after': '120', 'w:line': '360', 'w:lineRule': 'auto' }, children: [] }] },
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Spaced'] }] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.properties.spacing.before).toBe(120);
    expect(para.properties.spacing.after).toBe(120);
    expect(para.properties.spacing.line).toBe(360);
    expect(para.properties.spacing.lineRule).toBe('auto');
  });

  it('handles multiple paragraphs', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['P1'] }] },
              ],
            },
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['P2'] }] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.body.blocks).toHaveLength(2);
  });

  it('parses paragraph style reference', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                {
                  tag: 'w:pPr',
                  attrs: {},
                  children: [
                    { tag: 'w:pStyle', attrs: { 'w:val': 'Heading1' }, children: [] },
                  ],
                },
                {
                  tag: 'w:r',
                  attrs: {},
                  children: [
                    { tag: 'w:t', attrs: {}, children: ['Title'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.style).toBe('Heading1');
  });

  it('parses paragraph numbering', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                {
                  tag: 'w:pPr',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:numPr',
                      attrs: {},
                      children: [
                        { tag: 'w:ilvl', attrs: { 'w:val': '0' }, children: [] },
                        { tag: 'w:numId', attrs: { 'w:val': '1' }, children: [] },
                      ],
                    },
                  ],
                },
                {
                  tag: 'w:r',
                  attrs: {},
                  children: [
                    { tag: 'w:t', attrs: {}, children: ['List item'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.numbering).toBeDefined();
    expect(para.numbering.level).toBe(0);
    expect(para.numbering.numId).toBe('1');
  });
});
