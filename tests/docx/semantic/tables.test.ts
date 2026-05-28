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

describe('DOCX semantic - tables', () => {
  it('parses a simple table', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:tbl',
              attrs: {},
              children: [
                {
                  tag: 'w:tr',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['A1'] }] }] },
                      ],
                    },
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['B1'] }] }] },
                      ],
                    },
                  ],
                },
                {
                  tag: 'w:tr',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['A2'] }] }] },
                      ],
                    },
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['B2'] }] }] },
                      ],
                    },
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
    const table = semantic.body.blocks[0] as any;
    expect(table.type).toBe('table');
    expect(table.rows).toHaveLength(2);
    expect(table.rows[0].cells).toHaveLength(2);
    expect(table.rows[0].cells[0].blocks[0].runs[0].text).toBe('A1');
    expect(table.rows[1].cells[1].blocks[0].runs[0].text).toBe('B2');
  });

  it('parses table with borders', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:tbl',
              attrs: {},
              children: [
                {
                  tag: 'w:tblPr',
                  attrs: {},
                  children: [
                    { tag: 'w:tblW', attrs: { 'w:w': '5000', 'w:type': 'dxa' }, children: [] },
                    {
                      tag: 'w:tblBorders',
                      attrs: {},
                      children: [
                        { tag: 'w:top', attrs: { 'w:val': 'single', 'w:sz': '4', 'w:color': '000000' }, children: [] },
                        { tag: 'w:bottom', attrs: { 'w:val': 'single', 'w:sz': '4', 'w:color': '000000' }, children: [] },
                      ],
                    },
                  ],
                },
                {
                  tag: 'w:tr',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Cell'] }] }] },
                      ],
                    },
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
    const table = semantic.body.blocks[0] as any;
    expect(table.properties.width).toBe(5000);
    expect(table.properties.borders.top.style).toBe('single');
    expect(table.properties.borders.top.size).toBe(4);
    expect(table.properties.borders.top.color).toBe('000000');
  });

  it('parses table cell vertical align', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:tbl',
              attrs: {},
              children: [
                {
                  tag: 'w:tr',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        {
                          tag: 'w:tcPr',
                          attrs: {},
                          children: [
                            { tag: 'w:vAlign', attrs: { 'w:val': 'center' }, children: [] },
                          ],
                        },
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Centered'] }] }] },
                      ],
                    },
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
    const cell = (semantic.body.blocks[0] as any).rows[0].cells[0];
    expect(cell.properties.verticalAlign).toBe('center');
  });

  it('parses table cell merge', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:tbl',
              attrs: {},
              children: [
                {
                  tag: 'w:tr',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        {
                          tag: 'w:tcPr',
                          attrs: {},
                          children: [
                            { tag: 'w:gridSpan', attrs: { 'w:val': '2' }, children: [] },
                          ],
                        },
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Merged'] }] }] },
                      ],
                    },
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['C'] }] }] },
                      ],
                    },
                  ],
                },
                {
                  tag: 'w:tr',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        {
                          tag: 'w:tcPr',
                          attrs: {},
                          children: [
                            { tag: 'w:vMerge', attrs: { 'w:val': 'restart' }, children: [] },
                          ],
                        },
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['VMerge'] }] }] },
                      ],
                    },
                    {
                      tag: 'w:tc',
                      attrs: {},
                      children: [
                        {
                          tag: 'w:tcPr',
                          attrs: {},
                          children: [
                            { tag: 'w:vMerge', attrs: {}, children: [] },
                          ],
                        },
                        { tag: 'w:p', attrs: {}, children: [{ tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Continue'] }] }] },
                      ],
                    },
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
    const table = semantic.body.blocks[0] as any;
    expect(table.rows[0].cells[0].properties.gridSpan).toBe(2);
    expect(table.rows[0].cells[0].properties.horizontalMerge).toBe('restart');
    expect(table.rows[1].cells[0].properties.verticalMerge).toBe('restart');
    expect(table.rows[1].cells[1].properties.verticalMerge).toBe('continue');
  });
});
