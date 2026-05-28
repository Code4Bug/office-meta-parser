import { describe, it, expect } from 'vitest';
import { serializeWorkbook } from '../../../src/xlsx/serializers/workbook.js';
import type { XlsxWorkbook } from '../../../src/xlsx/types.js';

function makeWb(overrides: Partial<XlsxWorkbook> = {}): XlsxWorkbook {
  return {
    meta: {},
    sheets: [{ name: 'Sheet1', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] }],
    styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
    sharedStrings: [],
    ...overrides,
  };
}

describe('XLSX P0 serializer - workbook', () => {
  it('serializes defined names', () => {
    const xml = serializeWorkbook(makeWb({
      definedNames: [
        { name: 'SalesData', formula: 'Sheet1!$A$1:$D$100' },
        { name: 'TaxRate', formula: '0.08', localSheetId: 0 },
      ],
    }));

    expect(xml).toContain('<definedNames>');
    expect(xml).toContain('<definedName name="SalesData">Sheet1!$A$1:$D$100</definedName>');
    expect(xml).toContain('<definedName name="TaxRate" localSheetId="0">0.08</definedName>');
  });

  it('serializes hidden defined name', () => {
    const xml = serializeWorkbook(makeWb({
      definedNames: [
        { name: 'Hidden', formula: 'Sheet1!$A$1', hidden: true },
      ],
    }));

    expect(xml).toContain('hidden="1"');
  });

  it('serializes sheet state', () => {
    const xml = serializeWorkbook(makeWb({
      sheets: [
        { name: 'Visible', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
        { name: 'Hidden', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [], state: 'hidden' },
        { name: 'VeryHidden', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [], state: 'veryHidden' },
      ],
    }));

    expect(xml).not.toContain('state="visible"');
    expect(xml).toContain('state="hidden"');
    expect(xml).toContain('state="veryHidden"');
  });

  it('omits definedNames when empty', () => {
    const xml = serializeWorkbook(makeWb({ definedNames: [] }));
    expect(xml).not.toContain('<definedNames>');
  });
});
