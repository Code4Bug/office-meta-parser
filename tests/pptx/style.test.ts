import { describe, it, expect } from 'vitest';
import { parseShapeStyle } from '../../src/pptx/parsers/style.js';
import type { ParsedNode } from '../../src/core/types.js';

function makeSpPr(children: ParsedNode[]): ParsedNode {
  return { tag: 'p:spPr', attrs: {}, children };
}

describe('parseShapeStyle', () => {
  it('returns undefined for empty spPr', () => {
    const spPr = makeSpPr([]);
    expect(parseShapeStyle(spPr)).toBeUndefined();
  });

  it('parses solid fill', () => {
    const spPr = makeSpPr([
      {
        tag: 'a:solidFill',
        attrs: {},
        children: [{ tag: 'a:srgbClr', attrs: { val: 'FF0000' }, children: [] }],
      },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.fill?.type).toBe('solid');
    expect(style?.fill?.color).toBe('FF0000');
  });

  it('parses gradient fill', () => {
    const spPr = makeSpPr([
      { tag: 'a:gradFill', attrs: {}, children: [] },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.fill?.type).toBe('gradient');
  });

  it('parses no fill', () => {
    const spPr = makeSpPr([
      { tag: 'a:noFill', attrs: {}, children: [] },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.fill?.type).toBe('none');
  });

  it('parses border with width and color', () => {
    const spPr = makeSpPr([
      {
        tag: 'a:ln',
        attrs: { w: '25400' }, // 2pt in EMU
        children: [
          {
            tag: 'a:solidFill',
            attrs: {},
            children: [{ tag: 'a:srgbClr', attrs: { val: '0000FF' }, children: [] }],
          },
        ],
      },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.border?.width).toBe(2);
    expect(style?.border?.color).toBe('0000FF');
  });

  it('parses dashed border', () => {
    const spPr = makeSpPr([
      {
        tag: 'a:ln',
        attrs: { w: '12700' },
        children: [
          { tag: 'a:prstDash', attrs: { val: 'dash' }, children: [] },
        ],
      },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.border?.width).toBe(1);
    expect(style?.border?.style).toBe('dashed');
  });

  it('parses outer shadow', () => {
    const spPr = makeSpPr([
      {
        tag: 'a:effectLst',
        attrs: {},
        children: [
          {
            tag: 'a:outerShdw',
            attrs: { blurRad: '50800', dist: '38100', dir: '2700000' },
            children: [{ tag: 'a:srgbClr', attrs: { val: '000000' }, children: [] }],
          },
        ],
      },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.shadow?.type).toBe('outer');
    expect(style?.shadow?.blur).toBe(4);
    expect(style?.shadow?.color).toBe('000000');
  });

  it('parses scheme color', () => {
    const spPr = makeSpPr([
      {
        tag: 'a:solidFill',
        attrs: {},
        children: [{ tag: 'a:schemeClr', attrs: { val: 'accent1' }, children: [] }],
      },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.fill?.color).toBe('scheme:accent1');
  });

  it('parses combined style', () => {
    const spPr = makeSpPr([
      {
        tag: 'a:solidFill',
        attrs: {},
        children: [{ tag: 'a:srgbClr', attrs: { val: 'FFFF00' }, children: [] }],
      },
      {
        tag: 'a:ln',
        attrs: { w: '12700' },
        children: [
          {
            tag: 'a:solidFill',
            attrs: {},
            children: [{ tag: 'a:srgbClr', attrs: { val: '000000' }, children: [] }],
          },
        ],
      },
    ]);

    const style = parseShapeStyle(spPr);
    expect(style?.fill?.type).toBe('solid');
    expect(style?.fill?.color).toBe('FFFF00');
    expect(style?.border?.color).toBe('000000');
  });
});
