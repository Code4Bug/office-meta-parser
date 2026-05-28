import { describe, it, expect } from 'vitest';
import { parseParagraphProperties } from '../../../src/docx/parsers/paragraph.js';
import type { ParsedNode } from '../../../src/core/types.js';

function makePPr(children: ParsedNode[]): ParsedNode {
  return { tag: 'w:pPr', attrs: {}, children };
}

describe('DOCX parser - paragraph properties P0', () => {
  it('parses outlineLevel', () => {
    const props = parseParagraphProperties(makePPr([
      { tag: 'w:outlineLvl', attrs: { 'w:val': '2' }, children: [] },
    ]));
    expect(props.outlineLevel).toBe(2);
  });

  it('parses keepNext', () => {
    const props = parseParagraphProperties(makePPr([
      { tag: 'w:keepNext', attrs: {}, children: [] },
    ]));
    expect(props.keepNext).toBe(true);
  });

  it('parses keepLines', () => {
    const props = parseParagraphProperties(makePPr([
      { tag: 'w:keepLines', attrs: {}, children: [] },
    ]));
    expect(props.keepLines).toBe(true);
  });

  it('parses pageBreakBefore', () => {
    const props = parseParagraphProperties(makePPr([
      { tag: 'w:pageBreakBefore', attrs: {}, children: [] },
    ]));
    expect(props.pageBreakBefore).toBe(true);
  });

  it('parses paragraph shading', () => {
    const props = parseParagraphProperties(makePPr([
      { tag: 'w:shd', attrs: { 'w:fill': 'FFFF00', 'w:val': 'clear' }, children: [] },
    ]));
    expect(props.shading?.fill).toBe('FFFF00');
    expect(props.shading?.pattern).toBe('clear');
  });

  it('parses paragraph border', () => {
    const props = parseParagraphProperties(makePPr([
      {
        tag: 'w:pBdr',
        attrs: {},
        children: [
          { tag: 'w:top', attrs: { 'w:val': 'single', 'w:sz': '4', 'w:color': '000000' }, children: [] },
        ],
      },
    ]));
    expect(props.border?.top?.style).toBe('single');
    expect(props.border?.top?.size).toBe(4);
    expect(props.border?.top?.color).toBe('000000');
  });

  it('parses tab stops', () => {
    const props = parseParagraphProperties(makePPr([
      {
        tag: 'w:tabs',
        attrs: {},
        children: [
          { tag: 'w:tab', attrs: { 'w:val': 'right', 'w:pos': '4320', 'w:leader': 'dot' }, children: [] },
        ],
      },
    ]));
    expect(props.tabs).toHaveLength(1);
    expect(props.tabs![0].position).toBe(4320);
    expect(props.tabs![0].alignment).toBe('right');
    expect(props.tabs![0].leader).toBe('dot');
  });
});
