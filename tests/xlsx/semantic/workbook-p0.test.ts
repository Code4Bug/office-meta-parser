import { describe, it, expect } from 'vitest';
import { extractDefinedNames, extractTheme } from '../../../src/xlsx/parsers/workbook.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('XLSX P0 semantic - workbook features', () => {
  it('parses defined names', () => {
    const wbXml: ParsedNode = {
      tag: 'workbook', attrs: {}, children: [{
        tag: 'definedNames', attrs: {}, children: [
          { tag: 'definedName', attrs: { name: 'SalesData' }, children: ['Sheet1!$A$1:$D$100'] },
          { tag: 'definedName', attrs: { name: 'TaxRate', localSheetId: '0' }, children: ['0.08'] },
          { tag: 'definedName', attrs: { name: 'HiddenRange', hidden: '1' }, children: ['Sheet1!$F$1:$F$10'] },
        ],
      }],
    };

    const raw = makeRaw({ 'xl/workbook.xml': wbXml });
    const names = extractDefinedNames(raw);
    expect(names).toHaveLength(3);
    expect(names![0]).toEqual({ name: 'SalesData', formula: 'Sheet1!$A$1:$D$100' });
    expect(names![1]).toEqual({ name: 'TaxRate', formula: '0.08', localSheetId: 0 });
    expect(names![2]).toEqual({ name: 'HiddenRange', formula: 'Sheet1!$F$1:$F$10', hidden: true });
  });

  it('returns undefined when no defined names', () => {
    const wbXml: ParsedNode = { tag: 'workbook', attrs: {}, children: [] };
    const raw = makeRaw({ 'xl/workbook.xml': wbXml });
    expect(extractDefinedNames(raw)).toBeUndefined();
  });

  it('parses theme color scheme and fonts', () => {
    const themeXml: ParsedNode = {
      tag: 'a:theme', attrs: {}, children: [{
        tag: 'a:themeElements', attrs: {}, children: [
          {
            tag: 'a:clrScheme', attrs: { name: 'Office' }, children: [
              { tag: 'a:dk1', attrs: {}, children: [{ tag: 'a:sysClr', attrs: { lastClr: '000000' }, children: [] }] },
              { tag: 'a:lt1', attrs: {}, children: [{ tag: 'a:sysClr', attrs: { lastClr: 'FFFFFF' }, children: [] }] },
              { tag: 'a:accent1', attrs: {}, children: [{ tag: 'a:srgbClr', attrs: { val: '4472C4' }, children: [] }] },
            ],
          },
          {
            tag: 'a:fontScheme', attrs: { name: 'Office' }, children: [
              { tag: 'a:majorFont', attrs: {}, children: [{ tag: 'a:latin', attrs: { typeface: 'Calibri Light' }, children: [] }] },
              { tag: 'a:minorFont', attrs: {}, children: [{ tag: 'a:latin', attrs: { typeface: 'Calibri' }, children: [] }] },
            ],
          },
        ],
      }],
    };

    const raw = makeRaw({ 'xl/theme/theme1.xml': themeXml });
    const theme = extractTheme(raw);
    expect(theme).toBeDefined();
    expect(theme!.colorScheme).toBeDefined();
    expect(theme!.colorScheme!['dk1']).toBe('000000');
    expect(theme!.colorScheme!['accent1']).toBe('4472C4');
    expect(theme!.majorFont).toBe('Calibri Light');
    expect(theme!.minorFont).toBe('Calibri');
  });

  it('returns undefined when no theme', () => {
    const raw = makeRaw({});
    expect(extractTheme(raw)).toBeUndefined();
  });
});
