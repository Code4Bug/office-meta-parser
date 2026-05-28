import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/xlsx/serializer.js';
import type { XlsxWorkbook } from '../../../src/xlsx/types.js';

function makeWorkbook(sheets: any[]): XlsxWorkbook {
  return {
    meta: {},
    sheets,
    styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
    sharedStrings: [],
  };
}

describe('XLSX serializer - features', () => {
  it('serializes auto filter', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
        autoFilter: {
          ref: 'A1:C10',
          columns: [
            { colId: 0, filters: ['Apple', 'Banana'] },
          ],
        },
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('autoFilter');
    expect(result.worksheets[0]).toContain('ref="A1:C10"');
    expect(result.worksheets[0]).toContain('filterColumn');
    expect(result.worksheets[0]).toContain('colId="0"');
    expect(result.worksheets[0]).toContain('filters');
    expect(result.worksheets[0]).toContain('val="Apple"');
    expect(result.worksheets[0]).toContain('val="Banana"');
  });

  it('serializes data validations', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
        dataValidations: [
          {
            type: 'list',
            sqref: 'B2:B10',
            allowBlank: true,
            showErrorMessage: true,
            errorTitle: 'Invalid',
            error: 'Select from list',
            formula1: '"A,B,C"',
          },
        ],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('dataValidations');
    expect(result.worksheets[0]).toContain('dataValidation');
    expect(result.worksheets[0]).toContain('type="list"');
    expect(result.worksheets[0]).toContain('sqref="B2:B10"');
    expect(result.worksheets[0]).toContain('formula1');
    expect(result.worksheets[0]).toContain('&quot;A,B,C&quot;');
  });

  it('serializes conditional formats', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
        conditionalFormats: [
          {
            sqref: 'A1:A10',
            rules: [
              { type: 'cellIs', priority: 1, formula: ['100'] },
            ],
          },
        ],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('conditionalFormatting');
    expect(result.worksheets[0]).toContain('sqref="A1:A10"');
    expect(result.worksheets[0]).toContain('cfRule');
    expect(result.worksheets[0]).toContain('type="cellIs"');
    expect(result.worksheets[0]).toContain('priority="1"');
    expect(result.worksheets[0]).toContain('formula');
    expect(result.worksheets[0]).toContain('100');
  });

  it('serializes print area', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
        printArea: {
          fitToWidth: 1,
          fitToHeight: 1,
          pageMargins: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7, header: 0.3, footer: 0.3 },
        },
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.worksheets[0]).toContain('printOptions');
    expect(result.worksheets[0]).toContain('fitToWidth="1"');
    expect(result.worksheets[0]).toContain('fitToHeight="1"');
    expect(result.worksheets[0]).toContain('pageMargins');
    expect(result.worksheets[0]).toContain('top="0.75"');
  });

  it('serializes images', () => {
    const wb = makeWorkbook([
      {
        name: 'Sheet1',
        cells: [],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
        images: [
          {
            relationshipId: 'rId1',
            name: 'Image1',
            description: 'Test image',
            position: {
              from: { col: 0, row: 0, colOff: 0, rowOff: 0 },
              to: { col: 2, row: 5, colOff: 0, rowOff: 0 },
            },
          },
        ],
      },
    ]);

    const result = semanticToXml(wb);
    expect(result.drawings).toHaveLength(1);
    expect(result.drawings[0]).toContain('xdr:wsDr');
    expect(result.drawings[0]).toContain('xdr:twoCellAnchor');
    expect(result.drawings[0]).toContain('xdr:pic');
    expect(result.drawings[0]).toContain('a:blip');
    expect(result.drawings[0]).toContain('r:embed="rId1"');
    expect(result.worksheets[0]).toContain('drawing');
  });
});
