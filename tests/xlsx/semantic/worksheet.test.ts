import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/xlsx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('XLSX semantic - worksheet', () => {
  it('parses worksheet with cells', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [
            {
              tag: 'row',
              attrs: { r: '1' },
              children: [
                {
                  tag: 'c',
                  attrs: { r: 'A1', t: 'inlineStr' },
                  children: [
                    {
                      tag: 'is',
                      attrs: {},
                      children: [
                        { tag: 't', attrs: {}, children: ['Hello'] },
                      ],
                    },
                  ],
                },
                {
                  tag: 'c',
                  attrs: { r: 'B1' },
                  children: [
                    { tag: 'v', attrs: {}, children: ['42'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    // Should have one sheet (default)
    expect(semantic.sheets).toHaveLength(1);
    expect(semantic.sheets[0].cells).toHaveLength(1); // one row
    expect(semantic.sheets[0].cells[0]).toHaveLength(2); // two cells
    expect(semantic.sheets[0].cells[0][0].value).toBe('Hello');
    expect(semantic.sheets[0].cells[0][1].value).toBe(42);
  });

  it('parses merged cells', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [
            {
              tag: 'row',
              attrs: { r: '1' },
              children: [
                {
                  tag: 'c',
                  attrs: { r: 'A1', t: 'inlineStr' },
                  children: [
                    {
                      tag: 'is',
                      attrs: {},
                      children: [
                        { tag: 't', attrs: {}, children: ['Merged'] },
                      ],
                    },
                  ],
                },
                {
                  tag: 'c',
                  attrs: { r: 'B1' },
                  children: [],
                },
                {
                  tag: 'c',
                  attrs: { r: 'C1' },
                  children: [],
                },
              ],
            },
          ],
        },
        {
          tag: 'mergeCells',
          attrs: {},
          children: [
            { tag: 'mergeCell', attrs: { ref: 'A1:C1' }, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].mergedCells).toHaveLength(1);
    expect(semantic.sheets[0].mergedCells[0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
    });
  });

  it('parses multiple merged cells', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [],
        },
        {
          tag: 'mergeCells',
          attrs: {},
          children: [
            { tag: 'mergeCell', attrs: { ref: 'A1:B2' }, children: [] },
            { tag: 'mergeCell', attrs: { ref: 'D4:E5' }, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].mergedCells).toHaveLength(2);
    expect(semantic.sheets[0].mergedCells[0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
    expect(semantic.sheets[0].mergedCells[1]).toEqual({
      startRow: 3,
      startCol: 3,
      endRow: 4,
      endCol: 4,
    });
  });

  it('parses column widths', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'cols',
          attrs: {},
          children: [
            { tag: 'col', attrs: { min: '1', max: '1', width: '15', customWidth: '1' }, children: [] },
            { tag: 'col', attrs: { min: '2', max: '2', width: '20', customWidth: '1' }, children: [] },
          ],
        },
        {
          tag: 'sheetData',
          attrs: {},
          children: [],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].columnWidths).toEqual([15, 20]);
  });

  it('parses row heights', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [
            { tag: 'row', attrs: { r: '1', ht: '30', customHeight: '1' }, children: [] },
            { tag: 'row', attrs: { r: '2', ht: '45', customHeight: '1' }, children: [] },
            { tag: 'row', attrs: { r: '3' }, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].rowHeights).toEqual([30, 45]);
  });

  it('parses date cells', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [
            {
              tag: 'row',
              attrs: { r: '1' },
              children: [
                {
                  tag: 'c',
                  attrs: { r: 'A1', s: '14' }, // Style 14 is d/m/yyyy
                  children: [
                    { tag: 'v', attrs: {}, children: ['44927'] }, // 2023-01-01
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    const cell = semantic.sheets[0].cells[0][0];
    expect(cell.type).toBe('date');
    expect(cell.value).toBe('2023-01-01');
  });

  it('parses hyperlinks', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [
            {
              tag: 'row',
              attrs: { r: '1' },
              children: [
                {
                  tag: 'c',
                  attrs: { r: 'A1', t: 'inlineStr' },
                  children: [
                    {
                      tag: 'is',
                      attrs: {},
                      children: [
                        { tag: 't', attrs: {}, children: ['Click me'] },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
        {
          tag: 'hyperlinks',
          attrs: {},
          children: [
            {
              tag: 'hyperlink',
              attrs: { ref: 'A1', 'r:id': 'rId1', tooltip: 'Visit site' },
              children: [],
            },
          ],
        },
      ],
    };

    const rels = new Map<string, import('../../../src/core/types.js').Relationship[]>();
    rels.set('xl/worksheets/_rels/sheet1.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', target: 'https://example.com' },
    ]);

    const raw: RawDocument = {
      entries: [],
      rels,
      contentTypes: [],
      parts: new Map([['xl/worksheets/sheet1.xml', worksheetXml]]),
    };

    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].hyperlinks).toHaveLength(1);
    expect(semantic.sheets[0].hyperlinks[0].ref).toBe('A1');
    expect(semantic.sheets[0].hyperlinks[0].url).toBe('https://example.com');
    expect(semantic.sheets[0].hyperlinks[0].tooltip).toBe('Visit site');
  });

  it('parses error cells', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [
            {
              tag: 'row',
              attrs: { r: '1' },
              children: [
                {
                  tag: 'c',
                  attrs: { r: 'A1', t: 'e' },
                  children: [
                    { tag: 'v', attrs: {}, children: ['#REF!'] },
                  ],
                },
                {
                  tag: 'c',
                  attrs: { r: 'B1', t: 'e' },
                  children: [
                    { tag: 'v', attrs: {}, children: ['#VALUE!'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].cells[0][0].type).toBe('error');
    expect(semantic.sheets[0].cells[0][0].value).toBe('#REF!');
    expect(semantic.sheets[0].cells[0][1].type).toBe('error');
    expect(semantic.sheets[0].cells[0][1].value).toBe('#VALUE!');
  });
});
