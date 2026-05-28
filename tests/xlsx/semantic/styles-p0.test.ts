import { describe, it, expect } from 'vitest';
import { extractStyles } from '../../../src/xlsx/parsers/styles.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(stylesXml: ParsedNode): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map([['xl/styles.xml', stylesXml]]),
  };
}

function makeStyleSheet(children: ParsedNode[]): ParsedNode {
  return { tag: 'styleSheet', attrs: {}, children };
}

describe('XLSX P0 semantic - styles', () => {
  it('parses font vertAlign (superscript/subscript)', () => {
    const raw = makeRaw(makeStyleSheet([{
      tag: 'fonts', attrs: {}, children: [
        { tag: 'font', attrs: {}, children: [
          { tag: 'vertAlign', attrs: { val: 'superscript' }, children: [] },
          { tag: 'sz', attrs: { val: '12' }, children: [] },
        ]},
        { tag: 'font', attrs: {}, children: [
          { tag: 'vertAlign', attrs: { val: 'subscript' }, children: [] },
        ]},
      ],
    }, { tag: 'fills', attrs: {}, children: [] }, { tag: 'borders', attrs: {}, children: [] }, { tag: 'cellXfs', attrs: {}, children: [] }]));

    const styles = extractStyles(raw);
    expect(styles.fonts[0].vertAlign).toBe('superscript');
    expect(styles.fonts[1].vertAlign).toBe('subscript');
  });

  it('parses font scheme (major/minor)', () => {
    const raw = makeRaw(makeStyleSheet([{
      tag: 'fonts', attrs: {}, children: [
        { tag: 'font', attrs: {}, children: [
          { tag: 'scheme', attrs: { val: 'major' }, children: [] },
        ]},
        { tag: 'font', attrs: {}, children: [
          { tag: 'scheme', attrs: { val: 'minor' }, children: [] },
        ]},
      ],
    }, { tag: 'fills', attrs: {}, children: [] }, { tag: 'borders', attrs: {}, children: [] }, { tag: 'cellXfs', attrs: {}, children: [] }]));

    const styles = extractStyles(raw);
    expect(styles.fonts[0].scheme).toBe('major');
    expect(styles.fonts[1].scheme).toBe('minor');
  });

  it('parses gradient fill', () => {
    const raw = makeRaw(makeStyleSheet([{ tag: 'fonts', attrs: {}, children: [] }, {
      tag: 'fills', attrs: {}, children: [{
        tag: 'fill', attrs: {}, children: [{
          tag: 'gradientFill', attrs: { type: 'linear', degree: '90' }, children: [
            { tag: 'stop', attrs: { position: '0' }, children: [{ tag: 'color', attrs: { rgb: 'FF0000' }, children: [] }] },
            { tag: 'stop', attrs: { position: '1' }, children: [{ tag: 'color', attrs: { rgb: '0000FF' }, children: [] }] },
          ],
        }],
      }],
    }, { tag: 'borders', attrs: {}, children: [] }, { tag: 'cellXfs', attrs: {}, children: [] }]));

    const styles = extractStyles(raw);
    expect(styles.fills[0].gradientFill).toEqual({
      type: 'linear',
      degree: 90,
      stops: [
        { position: 0, color: 'FF0000' },
        { position: 1, color: '0000FF' },
      ],
    });
  });

  it('parses diagonal border', () => {
    const raw = makeRaw(makeStyleSheet([{ tag: 'fonts', attrs: {}, children: [] }, { tag: 'fills', attrs: {}, children: [] }, {
      tag: 'borders', attrs: {}, children: [{
        tag: 'border', attrs: { diagonalUp: '1', diagonalDown: '1' }, children: [
          { tag: 'diagonal', attrs: { style: 'thin' }, children: [{ tag: 'color', attrs: { rgb: 'FF0000' }, children: [] }] },
        ],
      }],
    }, { tag: 'cellXfs', attrs: {}, children: [] }]));

    const styles = extractStyles(raw);
    expect(styles.borders[0].diagonal).toEqual({ style: 'thin', color: 'FF0000' });
    expect(styles.borders[0].diagonalUp).toBe(true);
    expect(styles.borders[0].diagonalDown).toBe(true);
  });

  it('parses alignment indent, textRotation, shrinkToFit', () => {
    const raw = makeRaw(makeStyleSheet([{ tag: 'fonts', attrs: {}, children: [] }, { tag: 'fills', attrs: {}, children: [] }, { tag: 'borders', attrs: {}, children: [] }, {
      tag: 'cellXfs', attrs: {}, children: [{
        tag: 'xf', attrs: {}, children: [{
          tag: 'alignment', attrs: { horizontal: 'left', indent: '2', textRotation: '45', shrinkToFit: '1' }, children: [],
        }],
      }],
    }]));

    const styles = extractStyles(raw);
    expect(styles.cellStyles[0].alignment).toEqual({
      horizontal: 'left',
      indent: 2,
      textRotation: 45,
      shrinkToFit: true,
    });
  });
});
