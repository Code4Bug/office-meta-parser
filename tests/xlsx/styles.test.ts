import { describe, it, expect } from 'vitest';
import { extractStyles } from '../../src/xlsx/parsers/styles.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('extractStyles', () => {
  it('returns empty when no styles.xml', () => {
    const raw = makeRaw({});
    const styles = extractStyles(raw);
    expect(styles.fonts).toEqual([]);
    expect(styles.fills).toEqual([]);
    expect(styles.borders).toEqual([]);
    expect(styles.numberFormats).toEqual([]);
    expect(styles.cellStyles).toEqual([]);
  });

  it('parses number formats', () => {
    const stylesXml: ParsedNode = {
      tag: 'styleSheet',
      attrs: {},
      children: [
        {
          tag: 'numFmts',
          attrs: {},
          children: [
            { tag: 'numFmt', attrs: { numFmtId: '164', formatCode: 'yyyy-mm-dd' }, children: [] },
            { tag: 'numFmt', attrs: { numFmtId: '165', formatCode: '#,##0.00' }, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.numberFormats).toHaveLength(2);
    expect(styles.numberFormats[0]).toEqual({ id: '164', formatCode: 'yyyy-mm-dd' });
    expect(styles.numberFormats[1]).toEqual({ id: '165', formatCode: '#,##0.00' });
  });

  it('parses fonts', () => {
    const stylesXml: ParsedNode = {
      tag: 'styleSheet',
      attrs: {},
      children: [
        {
          tag: 'fonts',
          attrs: {},
          children: [
            {
              tag: 'font',
              attrs: {},
              children: [
                { tag: 'name', attrs: { val: 'Arial' }, children: [] },
                { tag: 'sz', attrs: { val: '12' }, children: [] },
                { tag: 'b', attrs: {}, children: [] },
                { tag: 'color', attrs: { rgb: 'FF000000' }, children: [] },
              ],
            },
            {
              tag: 'font',
              attrs: {},
              children: [
                { tag: 'name', attrs: { val: 'Calibri' }, children: [] },
                { tag: 'sz', attrs: { val: '11' }, children: [] },
                { tag: 'i', attrs: {}, children: [] },
                { tag: 'u', attrs: {}, children: [] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.fonts).toHaveLength(2);
    expect(styles.fonts[0].name).toBe('Arial');
    expect(styles.fonts[0].size).toBe(12);
    expect(styles.fonts[0].bold).toBe(true);
    expect(styles.fonts[0].color).toBe('FF000000');
    expect(styles.fonts[1].name).toBe('Calibri');
    expect(styles.fonts[1].italic).toBe(true);
    expect(styles.fonts[1].underline).toBe(true);
  });

  it('parses fills', () => {
    const stylesXml: ParsedNode = {
      tag: 'styleSheet',
      attrs: {},
      children: [
        {
          tag: 'fills',
          attrs: {},
          children: [
            {
              tag: 'fill',
              attrs: {},
              children: [
                {
                  tag: 'patternFill',
                  attrs: { patternType: 'solid' },
                  children: [
                    { tag: 'fgColor', attrs: { rgb: 'FFFFFF00' }, children: [] },
                    { tag: 'bgColor', attrs: { indexed: '64' }, children: [] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.fills).toHaveLength(1);
    expect(styles.fills[0].patternType).toBe('solid');
    expect(styles.fills[0].fgColor).toBe('FFFFFF00');
    expect(styles.fills[0].bgColor).toBe('indexed:64');
  });

  it('parses borders', () => {
    const stylesXml: ParsedNode = {
      tag: 'styleSheet',
      attrs: {},
      children: [
        {
          tag: 'borders',
          attrs: {},
          children: [
            {
              tag: 'border',
              attrs: {},
              children: [
                { tag: 'top', attrs: { style: 'thin' }, children: [{ tag: 'color', attrs: { rgb: 'FF000000' }, children: [] }] },
                { tag: 'bottom', attrs: { style: 'thin' }, children: [{ tag: 'color', attrs: { rgb: 'FF000000' }, children: [] }] },
                { tag: 'left', attrs: {}, children: [] },
                { tag: 'right', attrs: {}, children: [] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.borders).toHaveLength(1);
    expect(styles.borders[0].top?.style).toBe('thin');
    expect(styles.borders[0].top?.color).toBe('FF000000');
    expect(styles.borders[0].bottom?.style).toBe('thin');
    expect(styles.borders[0].left).toEqual({});
    expect(styles.borders[0].right).toEqual({});
  });

  it('parses cellXfs', () => {
    const stylesXml: ParsedNode = {
      tag: 'styleSheet',
      attrs: {},
      children: [
        {
          tag: 'cellXfs',
          attrs: {},
          children: [
            {
              tag: 'xf',
              attrs: { xfId: '0', fontId: '0', fillId: '0', borderId: '0' },
              children: [],
            },
            {
              tag: 'xf',
              attrs: { xfId: '1', fontId: '1', numFmtId: '164' },
              children: [
                { tag: 'alignment', attrs: { horizontal: 'center', vertical: 'center', wrapText: '1' }, children: [] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.cellStyles).toHaveLength(2);
    expect(styles.cellStyles[0].id).toBe('0');
    expect(styles.cellStyles[1].id).toBe('1');
    expect(styles.cellStyles[1].numberFormat).toBe('164');
    expect(styles.cellStyles[1].alignment?.horizontal).toBe('center');
    expect(styles.cellStyles[1].alignment?.vertical).toBe('center');
    expect(styles.cellStyles[1].alignment?.wrapText).toBe(true);
  });

  it('parses theme colors', () => {
    const stylesXml: ParsedNode = {
      tag: 'styleSheet',
      attrs: {},
      children: [
        {
          tag: 'fonts',
          attrs: {},
          children: [
            {
              tag: 'font',
              attrs: {},
              children: [
                { tag: 'color', attrs: { theme: '1' }, children: [] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'xl/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.fonts[0].color).toBe('theme:1');
  });
});
