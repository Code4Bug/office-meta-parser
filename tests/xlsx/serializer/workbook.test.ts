import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/xlsx/serializer.js';
import type { XlsxWorkbook } from '../../../src/xlsx/types.js';

describe('XLSX serializer - workbook', () => {
  it('serializes a workbook with sheets', () => {
    const wb: XlsxWorkbook = {
      meta: {},
      sheets: [
        { name: 'Sheet1', cells: [], mergedCells: [], columnWidths: [], rowHeights: [] },
      ],
      styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
      sharedStrings: [],
    };

    const result = semanticToXml(wb);
    expect(result.workbook).toContain('workbook');
    expect(result.workbook).toContain('Sheet1');
  });

  it('serializes shared strings', () => {
    const wb: XlsxWorkbook = {
      meta: {},
      sheets: [],
      styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
      sharedStrings: [{ text: 'Hello' }, { text: 'World' }],
    };

    const result = semanticToXml(wb);
    expect(result.sharedStrings).toContain('sst');
    expect(result.sharedStrings).toContain('Hello');
    expect(result.sharedStrings).toContain('World');
  });

  it('generates relationship files', () => {
    const wb: XlsxWorkbook = {
      meta: {},
      sheets: [
        { name: 'Sheet1', cells: [], mergedCells: [], columnWidths: [], rowHeights: [] },
      ],
      styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
      sharedStrings: [],
    };

    const result = semanticToXml(wb);
    expect(result.rels).toContain('Relationships');
    expect(result.sheetRels).toBeDefined();
  });
});
