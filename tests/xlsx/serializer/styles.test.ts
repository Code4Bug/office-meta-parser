import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/xlsx/serializer.js';
import type { XlsxWorkbook } from '../../../src/xlsx/types.js';

describe('XLSX serializer - styles', () => {
  it('serializes styles', () => {
    const wb: XlsxWorkbook = {
      meta: {},
      sheets: [],
      styles: {
        cellStyles: [
          {
            id: '0',
            alignment: { horizontal: 'center', vertical: 'center' },
          },
        ],
        fonts: [
          { name: 'Arial', size: 11, bold: true },
        ],
        fills: [
          { patternType: 'solid', fgColor: 'FFFFFF00' },
        ],
        borders: [
          { top: { style: 'thin', color: '000000' } },
        ],
        numberFormats: [
          { id: '164', formatCode: '#,##0.00' },
        ],
      },
      sharedStrings: [],
    };

    const result = semanticToXml(wb);
    expect(result.styles).toContain('styleSheet');
    expect(result.styles).toContain('fonts');
    expect(result.styles).toContain('Arial');
    expect(result.styles).toContain('fills');
    expect(result.styles).toContain('patternFill');
    expect(result.styles).toContain('borders');
    expect(result.styles).toContain('numFmts');
    expect(result.styles).toContain('cellXfs');
    expect(result.styles).toContain('alignment');
  });
});
