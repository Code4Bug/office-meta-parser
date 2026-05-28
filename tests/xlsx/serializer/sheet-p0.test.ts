import { describe, it, expect } from 'vitest';
import { serializeWorksheet } from '../../../src/xlsx/serializers/worksheet.js';
import type { Sheet } from '../../../src/xlsx/types.js';

function makeSheet(overrides: Partial<Sheet> = {}): Sheet {
  return {
    name: 'Sheet1',
    cells: [],
    mergedCells: [],
    columnWidths: [],
    rowHeights: [],
    hyperlinks: [],
    ...overrides,
  };
}

describe('XLSX P0 serializer - sheet features', () => {
  it('serializes frozen panes', () => {
    const xml = serializeWorksheet(makeSheet({
      frozenPanes: { xSplit: 1, ySplit: 2, topLeftCell: 'B3' },
    }), 0, []);

    expect(xml).toContain('<sheetViews>');
    expect(xml).toContain('<pane');
    expect(xml).toContain('xSplit="1"');
    expect(xml).toContain('ySplit="2"');
    expect(xml).toContain('topLeftCell="B3"');
    expect(xml).toContain('state="frozen"');
  });

  it('serializes zoom scale', () => {
    const xml = serializeWorksheet(makeSheet({ zoomScale: 150 }), 0, []);
    expect(xml).toContain('zoomScale="150"');
  });

  it('serializes active cell and selections', () => {
    const xml = serializeWorksheet(makeSheet({
      activeCell: 'C5',
      selections: [{ pane: 'bottomRight', activeCell: 'C5', sqref: 'C5:E10' }],
    }), 0, []);

    expect(xml).toContain('activeCell="C5"');
    expect(xml).toContain('sqref="C5:E10"');
  });

  it('serializes tab color', () => {
    const xml = serializeWorksheet(makeSheet({ tabColor: 'FFFF0000' }), 0, []);
    expect(xml).toContain('<sheetPr>');
    expect(xml).toContain('<tabColor');
    expect(xml).toContain('rgb="FFFF0000"');
  });

  it('serializes default row height and column width', () => {
    const xml = serializeWorksheet(makeSheet({
      defaultRowHeight: 20,
      defaultColWidth: 12,
    }), 0, []);

    expect(xml).toContain('<sheetFormatPr');
    expect(xml).toContain('defaultRowHeight="20"');
    expect(xml).toContain('defaultColWidth="12"');
  });

  it('serializes row/col outline groups', () => {
    const xml = serializeWorksheet(makeSheet({
      columnWidths: [10],
      colGroups: [{ level: 2, collapsed: true }],
      rowGroups: [{ level: 1, collapsed: false }],
    }), 0, []);

    expect(xml).toContain('outlineLevel="2"');
    expect(xml).toContain('collapsed="1"');
  });

  it('serializes page setup', () => {
    const xml = serializeWorksheet(makeSheet({
      pageSetup: { orientation: 'landscape', paperSize: 9, scale: 80 },
    }), 0, []);

    expect(xml).toContain('<pageSetup');
    expect(xml).toContain('orientation="landscape"');
    expect(xml).toContain('paperSize="9"');
    expect(xml).toContain('scale="80"');
  });

  it('serializes header/footer', () => {
    const xml = serializeWorksheet(makeSheet({
      headerFooter: {
        differentFirst: true,
        oddHeader: '&LPage &P',
        oddFooter: '&RConfidential',
        firstHeader: 'Title Page',
      },
    }), 0, []);

    expect(xml).toContain('<headerFooter');
    expect(xml).toContain('differentFirst="1"');
    expect(xml).toContain('<oddHeader>');
    expect(xml).toContain('&amp;LPage &amp;P');
    expect(xml).toContain('<firstHeader>Title Page</firstHeader>');
  });

  it('serializes table parts', () => {
    const xml = serializeWorksheet(makeSheet({
      tables: [{ id: 1, name: 'Sales', displayName: 'Sales', ref: 'A1:C10', columns: [] }],
    }), 0, []);

    expect(xml).toContain('<tableParts');
    expect(xml).toContain('<tablePart');
    expect(xml).toContain('r:id="rId1"');
  });

  it('round-trips frozen panes through parse and serialize', () => {
    const original = makeSheet({
      frozenPanes: { xSplit: 2, ySplit: 1, topLeftCell: 'C2' },
      zoomScale: 120,
      tabColor: 'FF00FF00',
      defaultRowHeight: 18,
    });

    const xml = serializeWorksheet(original, 0, []);

    expect(xml).toContain('xSplit="2"');
    expect(xml).toContain('ySplit="1"');
    expect(xml).toContain('topLeftCell="C2"');
    expect(xml).toContain('zoomScale="120"');
    expect(xml).toContain('rgb="FF00FF00"');
    expect(xml).toContain('defaultRowHeight="18"');
  });
});
