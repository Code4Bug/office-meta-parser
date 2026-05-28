import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/xlsx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';
import type { Relationship } from '../../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>, rels?: Map<string, Relationship[]>): RawDocument {
  return {
    entries: [],
    rels: rels || new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

function makeSheetXml(children: ParsedNode[]): ParsedNode {
  return { tag: 'worksheet', attrs: {}, children };
}

describe('XLSX P0 semantic - sheet features', () => {
  it('parses frozen panes', () => {
    const ws = makeSheetXml([{
      tag: 'sheetViews', attrs: {}, children: [{
        tag: 'sheetView', attrs: { workbookViewId: '0' }, children: [{
          tag: 'pane', attrs: { xSplit: '1', ySplit: '2', topLeftCell: 'B3', activePane: 'bottomRight', state: 'frozen' }, children: [],
        }],
      }],
    }, { tag: 'sheetData', attrs: {}, children: [] }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].frozenPanes).toEqual({ xSplit: 1, ySplit: 2, topLeftCell: 'B3' });
  });

  it('parses zoom scale', () => {
    const ws = makeSheetXml([{
      tag: 'sheetViews', attrs: {}, children: [{
        tag: 'sheetView', attrs: { workbookViewId: '0', zoomScale: '150' }, children: [],
      }],
    }, { tag: 'sheetData', attrs: {}, children: [] }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].zoomScale).toBe(150);
  });

  it('parses active cell and selections', () => {
    const ws = makeSheetXml([{
      tag: 'sheetViews', attrs: {}, children: [{
        tag: 'sheetView', attrs: { workbookViewId: '0' }, children: [{
          tag: 'selection', attrs: { activeCell: 'C5', sqref: 'C5:E10' }, children: [],
        }],
      }],
    }, { tag: 'sheetData', attrs: {}, children: [] }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].activeCell).toBe('C5');
    expect(sem.sheets[0].selections).toHaveLength(1);
    expect(sem.sheets[0].selections![0].activeCell).toBe('C5');
    expect(sem.sheets[0].selections![0].sqref).toBe('C5:E10');
  });

  it('parses sheet hidden state', () => {
    const wbXml: ParsedNode = {
      tag: 'workbook', attrs: {}, children: [{
        tag: 'sheets', attrs: {}, children: [
          { tag: 'sheet', attrs: { name: 'Sheet1', 'r:id': 'rId1' }, children: [] },
          { tag: 'sheet', attrs: { name: 'Sheet2', 'r:id': 'rId2', state: 'hidden' }, children: [] },
          { tag: 'sheet', attrs: { name: 'Sheet3', 'r:id': 'rId3', state: 'veryHidden' }, children: [] },
        ],
      }],
    };
    const ws1 = makeSheetXml([{ tag: 'sheetData', attrs: {}, children: [] }]);
    const ws2 = makeSheetXml([{ tag: 'sheetData', attrs: {}, children: [] }]);
    const ws3 = makeSheetXml([{ tag: 'sheetData', attrs: {}, children: [] }]);

    const rels = new Map<string, Relationship[]>();
    rels.set('xl/_rels/workbook.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', target: 'worksheets/sheet1.xml' },
      { id: 'rId2', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', target: 'worksheets/sheet2.xml' },
      { id: 'rId3', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', target: 'worksheets/sheet3.xml' },
    ]);

    const raw = makeRaw({
      'xl/workbook.xml': wbXml,
      'xl/worksheets/sheet1.xml': ws1,
      'xl/worksheets/sheet2.xml': ws2,
      'xl/worksheets/sheet3.xml': ws3,
    }, rels);

    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].state).toBeUndefined();
    expect(sem.sheets[1].state).toBe('hidden');
    expect(sem.sheets[2].state).toBe('veryHidden');
  });

  it('parses tab color', () => {
    const ws = makeSheetXml([{
      tag: 'sheetPr', attrs: {}, children: [{
        tag: 'tabColor', attrs: { rgb: 'FFFF0000' }, children: [],
      }],
    }, { tag: 'sheetData', attrs: {}, children: [] }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].tabColor).toBe('FFFF0000');
  });

  it('parses default row height and column width', () => {
    const ws = makeSheetXml([{
      tag: 'sheetFormatPr', attrs: { defaultRowHeight: '20', defaultColWidth: '12' }, children: [],
    }, { tag: 'sheetData', attrs: {}, children: [] }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].defaultRowHeight).toBe(20);
    expect(sem.sheets[0].defaultColWidth).toBe(12);
  });

  it('parses row/col outline groups', () => {
    const ws = makeSheetXml([{
      tag: 'cols', attrs: {}, children: [
        { tag: 'col', attrs: { min: '1', max: '1', width: '10', outlineLevel: '2', collapsed: '1' }, children: [] },
      ],
    }, {
      tag: 'sheetData', attrs: {}, children: [
        { tag: 'row', attrs: { r: '1', outlineLevel: '1' }, children: [] },
        { tag: 'row', attrs: { r: '2', outlineLevel: '1', collapsed: '1' }, children: [] },
      ],
    }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].colGroups).toHaveLength(1);
    expect(sem.sheets[0].colGroups![0]).toEqual({ level: 2, collapsed: true });
    expect(sem.sheets[0].rowGroups).toHaveLength(2);
    expect(sem.sheets[0].rowGroups![0]).toEqual({ level: 1, collapsed: false });
    expect(sem.sheets[0].rowGroups![1]).toEqual({ level: 1, collapsed: true });
  });

  it('parses page setup', () => {
    const ws = makeSheetXml([{ tag: 'sheetData', attrs: {}, children: [] }, {
      tag: 'pageSetup', attrs: { orientation: 'landscape', paperSize: '9', scale: '80', fitToWidth: '1', fitToHeight: '0' }, children: [],
    }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].pageSetup).toEqual({
      orientation: 'landscape', paperSize: 9, scale: 80, fitToWidth: 1, fitToHeight: 0,
    });
  });

  it('parses header/footer', () => {
    const ws = makeSheetXml([{ tag: 'sheetData', attrs: {}, children: [] }, {
      tag: 'headerFooter', attrs: { differentFirst: '1' }, children: [
        { tag: 'oddHeader', attrs: {}, children: ['&LPage &P'] },
        { tag: 'oddFooter', attrs: {}, children: ['&RConfidential'] },
        { tag: 'firstHeader', attrs: {}, children: ['Title Page'] },
      ],
    }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].headerFooter).toEqual({
      differentFirst: true,
      oddHeader: '&LPage &P',
      oddFooter: '&RConfidential',
      firstHeader: 'Title Page',
    });
  });

  it('parses array/shared formula types', () => {
    const ws = makeSheetXml([{
      tag: 'sheetData', attrs: {}, children: [{
        tag: 'row', attrs: { r: '1' }, children: [
          { tag: 'c', attrs: { r: 'A1' }, children: [{ tag: 'f', attrs: { t: 'array', ref: 'A1:A3' }, children: ['{=SUM(B1:B3*C1:C3)}'] }, { tag: 'v', attrs: {}, children: ['10'] }] },
          { tag: 'c', attrs: { r: 'A2' }, children: [{ tag: 'f', attrs: { t: 'shared', si: '0' }, children: ['B2+C2'] }, { tag: 'v', attrs: {}, children: ['20'] }] },
        ],
      }],
    }]);

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].cells[0][0].formulaType).toBe('array');
    expect(sem.sheets[0].cells[0][0].formulaRef).toBe('A1:A3');
    expect(sem.sheets[0].cells[0][1].formulaType).toBe('shared');
    expect(sem.sheets[0].cells[0][1].sharedFormulaIndex).toBe(0);
  });

  it('parses comments', () => {
    const ws = makeSheetXml([{ tag: 'sheetData', attrs: {}, children: [] }]);
    const commentsXml: ParsedNode = {
      tag: 'comments', attrs: {}, children: [
        { tag: 'authors', attrs: {}, children: [{ tag: 'author', attrs: {}, children: ['John'] }] },
        {
          tag: 'commentList', attrs: {}, children: [{
            tag: 'comment', attrs: { ref: 'A1', authorId: '0' }, children: [
              { tag: 'text', attrs: {}, children: ['This is a comment'] },
            ],
          }],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': ws, 'xl/comments1.xml': commentsXml });
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].comments).toHaveLength(1);
    expect(sem.sheets[0].comments![0].ref).toBe('A1');
    expect(sem.sheets[0].comments![0].text).toBe('This is a comment');
  });

  it('parses Excel tables', () => {
    const ws = makeSheetXml([{ tag: 'sheetData', attrs: {}, children: [] }]);
    const tableXml: ParsedNode = {
      tag: 'table', attrs: { id: '1', name: 'SalesTable', displayName: 'SalesTable', ref: 'A1:C10', headerRowCount: '1' }, children: [{
        tag: 'tableColumns', attrs: {}, children: [
          { tag: 'tableColumn', attrs: { id: '1', name: 'Product' }, children: [] },
          { tag: 'tableColumn', attrs: { id: '2', name: 'Amount' }, children: [] },
          { tag: 'tableColumn', attrs: { id: '3', name: 'Total', totalsRowFunction: 'sum' }, children: [] },
        ],
      }],
    };

    const wbXml: ParsedNode = {
      tag: 'workbook', attrs: {}, children: [{
        tag: 'sheets', attrs: {}, children: [
          { tag: 'sheet', attrs: { name: 'Sheet1', 'r:id': 'rId1' }, children: [] },
        ],
      }],
    };

    const rels = new Map<string, Relationship[]>();
    rels.set('xl/_rels/workbook.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', target: 'worksheets/sheet1.xml' },
    ]);
    rels.set('xl/worksheets/_rels/sheet1.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', target: '../tables/table1.xml' },
    ]);

    const raw = makeRaw({
      'xl/workbook.xml': wbXml,
      'xl/worksheets/sheet1.xml': ws,
      'xl/tables/table1.xml': tableXml,
    }, rels);
    const sem = rawToSemantic(raw);
    expect(sem.sheets[0].tables).toHaveLength(1);
    expect(sem.sheets[0].tables![0].name).toBe('SalesTable');
    expect(sem.sheets[0].tables![0].columns).toHaveLength(3);
    expect(sem.sheets[0].tables![0].columns[2].totalsRowFunction).toBe('sum');
  });
});
