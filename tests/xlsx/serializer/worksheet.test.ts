import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/xlsx/serializer.js';
import type { XlsxWorkbook, SharedStringEntry } from '../../../src/xlsx/types.js';

function makeWorkbook(sheets: any[], sharedStrings: SharedStringEntry[] = []): XlsxWorkbook {
  return {
    meta: {},
    sheets,
    styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
    sharedStrings,
  };
}

describe('XLSX serializer - worksheet', () => {
  it('serializes worksheet with cells', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [
          [
            { value: 'Hello', type: 'string' },
            { value: 42, type: 'number' },
          ],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
      },
    ], [{ text: 'Hello' }]);

    const result = semanticToXml(wb);
    expect(result.worksheets).toHaveLength(1);
    expect(result.worksheets[0]).toContain('sheetData');
    expect(result.worksheets[0]).toContain('Hello');
    expect(result.worksheets[0]).toContain('42');
  });

  it('serializes merged cells', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [
          [{ value: 'A1', type: 'string' }, { value: 'B1', type: 'string' }],
          [{ value: 'A2', type: 'string' }, { value: 'B2', type: 'string' }],
        ],
        mergedCells: [
          { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
        ],
        columnWidths: [],
        rowHeights: [],
      },
    ], [{ text: 'A1' }, { text: 'B1' }, { text: 'A2' }, { text: 'B2' }]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('mergeCells');
    expect(result.worksheets[0]).toContain('mergeCell');
    expect(result.worksheets[0]).toContain('A1:B1');
  });

  it('serializes column widths', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [],
        mergedCells: [],
        columnWidths: [10, 20, 15],
        rowHeights: [],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('cols');
    expect(result.worksheets[0]).toContain('col');
    expect(result.worksheets[0]).toContain('width="10"');
    expect(result.worksheets[0]).toContain('width="20"');
  });

  it('serializes row heights', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [
          [{ value: 'A1', type: 'string' }],
          [{ value: 'A2', type: 'string' }],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [30, 50],
      },
    ], [{ text: 'A1' }, { text: 'A2' }]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('ht="30"');
    expect(result.worksheets[0]).toContain('ht="50"');
    expect(result.worksheets[0]).toContain('customHeight="1"');
  });

  it('serializes inline strings', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [
          [{ value: 'Hello', type: 'string' }],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('t="inlineStr"');
    expect(result.worksheets[0]).toContain('<is><t>Hello</t></is>');
  });

  it('serializes numbers', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [
          [{ value: 42, type: 'number' }, { value: 3.14, type: 'number' }],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('<v>42</v>');
    expect(result.worksheets[0]).toContain('<v>3.14</v>');
  });

  it('serializes booleans', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [
          [{ value: true, type: 'boolean' }, { value: false, type: 'boolean' }],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('t="b"');
    expect(result.worksheets[0]).toContain('<v>1</v>');
    expect(result.worksheets[0]).toContain('<v>0</v>');
  });

  it('serializes formulas', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [
          [{ value: 10, type: 'number' }, { value: 20, type: 'number' }, { value: 30, type: 'number', formula: 'A1+B1' }],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('<f>A1+B1</f>');
    expect(result.worksheets[0]).toContain('<v>30</v>');
  });
});
