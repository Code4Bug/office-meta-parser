import { describe, it, expect } from 'vitest';
import type {
  XlsxWorkbook,
  Sheet,
  Cell,
  CellStyle,
  MergedCell,
} from '../../src/xlsx/types.js';

describe('XLSX types', () => {
  it('XlsxWorkbook structure', () => {
    const wb: XlsxWorkbook = {
      meta: { title: 'Test' },
      sheets: [],
      styles: { cellStyles: [] },
      sharedStrings: [],
    };
    expect(wb.meta.title).toBe('Test');
    expect(wb.sheets).toEqual([]);
  });

  it('Sheet structure', () => {
    const sheet: Sheet = {
      name: 'Sheet1',
      cells: [],
      mergedCells: [],
      columnWidths: [],
      rowHeights: [],
    };
    expect(sheet.name).toBe('Sheet1');
  });

  it('Cell with value', () => {
    const cell: Cell = {
      value: 'Hello',
      type: 'string',
    };
    expect(cell.value).toBe('Hello');
    expect(cell.type).toBe('string');
  });

  it('Cell with number', () => {
    const cell: Cell = {
      value: 42,
      type: 'number',
    };
    expect(cell.value).toBe(42);
  });

  it('Cell with formula', () => {
    const cell: Cell = {
      value: null,
      formula: 'SUM(A1:A10)',
      type: 'formula',
    };
    expect(cell.formula).toBe('SUM(A1:A10)');
  });

  it('CellStyle structure', () => {
    const style: CellStyle = {
      id: '1',
      font: { bold: true, color: 'FF0000' },
      fill: { bgColor: 'FFFF00' },
      border: { top: { style: 'thin' } },
    };
    expect(style.font?.bold).toBe(true);
  });

  it('MergedCell structure', () => {
    const merged: MergedCell = {
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 2,
    };
    expect(merged.startRow).toBe(0);
    expect(merged.endCol).toBe(2);
  });
});
