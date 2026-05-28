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

describe('XLSX semantic - features', () => {
  it('parses auto filter', () => {
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
          tag: 'autoFilter',
          attrs: { ref: 'A1:C10' },
          children: [
            {
              tag: 'filterColumn',
              attrs: { colId: '0' },
              children: [
                {
                  tag: 'filters',
                  attrs: {},
                  children: [
                    { tag: 'filter', attrs: { val: 'Apple' }, children: [] },
                    { tag: 'filter', attrs: { val: 'Banana' }, children: [] },
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
    expect(semantic.sheets[0].autoFilter).toBeDefined();
    expect(semantic.sheets[0].autoFilter!.ref).toBe('A1:C10');
    expect(semantic.sheets[0].autoFilter!.columns).toHaveLength(1);
    expect(semantic.sheets[0].autoFilter!.columns[0].colId).toBe(0);
    expect(semantic.sheets[0].autoFilter!.columns[0].filters).toEqual(['Apple', 'Banana']);
  });

  it('parses data validations', () => {
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
          tag: 'dataValidations',
          attrs: {},
          children: [
            {
              tag: 'dataValidation',
              attrs: {
                type: 'list',
                allowBlank: '1',
                showErrorMessage: '1',
                errorTitle: 'Invalid Input',
                error: 'Please select from the list',
                sqref: 'B2:B10',
              },
              children: [
                { tag: 'formula1', attrs: {}, children: ['"Apple,Banana,Cherry"'] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].dataValidations).toHaveLength(1);
    expect(semantic.sheets[0].dataValidations![0].type).toBe('list');
    expect(semantic.sheets[0].dataValidations![0].sqref).toBe('B2:B10');
    expect(semantic.sheets[0].dataValidations![0].formula1).toBe('"Apple,Banana,Cherry"');
    expect(semantic.sheets[0].dataValidations![0].allowBlank).toBe(true);
    expect(semantic.sheets[0].dataValidations![0].errorTitle).toBe('Invalid Input');
  });

  it('parses conditional formats', () => {
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
          tag: 'conditionalFormatting',
          attrs: { sqref: 'A1:A10' },
          children: [
            {
              tag: 'cfRule',
              attrs: { type: 'cellIs', priority: '1', operator: 'greaterThan' },
              children: [
                { tag: 'formula', attrs: {}, children: ['100'] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].conditionalFormats).toHaveLength(1);
    expect(semantic.sheets[0].conditionalFormats![0].sqref).toBe('A1:A10');
    expect(semantic.sheets[0].conditionalFormats![0].rules).toHaveLength(1);
    expect(semantic.sheets[0].conditionalFormats![0].rules[0].type).toBe('cellIs');
    expect(semantic.sheets[0].conditionalFormats![0].rules[0].formula).toEqual(['100']);
  });

  it('parses print area', () => {
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
          tag: 'printOptions',
          attrs: { fitToWidth: '1', fitToHeight: '1' },
          children: [],
        },
        {
          tag: 'pageMargins',
          attrs: { top: '0.75', right: '0.7', bottom: '0.75', left: '0.7', header: '0.3', footer: '0.3' },
          children: [],
        },
      ],
    };

    const raw = makeRaw({ 'xl/worksheets/sheet1.xml': worksheetXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].printArea).toBeDefined();
    expect(semantic.sheets[0].printArea!.fitToWidth).toBe(1);
    expect(semantic.sheets[0].printArea!.fitToHeight).toBe(1);
    expect(semantic.sheets[0].printArea!.pageMargins).toBeDefined();
    expect(semantic.sheets[0].printArea!.pageMargins!.top).toBe(0.75);
  });

  it('parses images from drawing', () => {
    const worksheetXml: ParsedNode = {
      tag: 'worksheet',
      attrs: {},
      children: [
        {
          tag: 'sheetData',
          attrs: {},
          children: [],
        },
      ],
    };

    const drawingXml: ParsedNode = {
      tag: 'xdr:wsDr',
      attrs: {},
      children: [
        {
          tag: 'xdr:twoCellAnchor',
          attrs: {},
          children: [
            {
              tag: 'xdr:from',
              attrs: {},
              children: [
                { tag: 'xdr:col', attrs: {}, children: ['0'] },
                { tag: 'xdr:colOff', attrs: {}, children: ['0'] },
                { tag: 'xdr:row', attrs: {}, children: ['0'] },
                { tag: 'xdr:rowOff', attrs: {}, children: ['0'] },
              ],
            },
            {
              tag: 'xdr:to',
              attrs: {},
              children: [
                { tag: 'xdr:col', attrs: {}, children: ['2'] },
                { tag: 'xdr:colOff', attrs: {}, children: ['0'] },
                { tag: 'xdr:row', attrs: {}, children: ['5'] },
                { tag: 'xdr:rowOff', attrs: {}, children: ['0'] },
              ],
            },
            {
              tag: 'xdr:pic',
              attrs: {},
              children: [
                {
                  tag: 'xdr:nvPicPr',
                  attrs: {},
                  children: [
                    { tag: 'xdr:cNvPr', attrs: { id: '1', name: 'Image1', descr: 'Test image' }, children: [] },
                    { tag: 'xdr:cNvPicPr', attrs: {}, children: [] },
                  ],
                },
                {
                  tag: 'xdr:blipFill',
                  attrs: {},
                  children: [
                    { tag: 'a:blip', attrs: { 'r:embed': 'rId1' }, children: [] },
                  ],
                },
                { tag: 'xdr:spPr', attrs: {}, children: [] },
              ],
            },
            { tag: 'xdr:clientData', attrs: {}, children: [] },
          ],
        },
      ],
    };

    const rels = new Map<string, import('../../../src/core/types.js').Relationship[]>();
    rels.set('xl/worksheets/_rels/sheet1.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', target: '../drawings/drawing1.xml' },
    ]);

    const raw: RawDocument = {
      entries: [],
      rels,
      contentTypes: [],
      parts: new Map([
        ['xl/worksheets/sheet1.xml', worksheetXml],
        ['xl/drawings/drawing1.xml', drawingXml],
      ]),
    };

    const semantic = rawToSemantic(raw);
    expect(semantic.sheets[0].images).toHaveLength(1);
    expect(semantic.sheets[0].images![0].relationshipId).toBe('rId1');
    expect(semantic.sheets[0].images![0].name).toBe('Image1');
    expect(semantic.sheets[0].images![0].description).toBe('Test image');
    expect(semantic.sheets[0].images![0].position.from.col).toBe(0);
    expect(semantic.sheets[0].images![0].position.from.row).toBe(0);
    expect(semantic.sheets[0].images![0].position.to.col).toBe(2);
    expect(semantic.sheets[0].images![0].position.to.row).toBe(5);
  });
});
