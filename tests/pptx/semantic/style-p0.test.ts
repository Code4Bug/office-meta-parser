import { describe, it, expect } from 'vitest';
import { parseShapeStyle, parseGradientFill, parseColorValue } from '../../../src/pptx/parsers/style.js';
import type { ParsedNode } from '../../../src/core/types.js';

function makeNode(tag: string, attrs: Record<string, string> = {}, children: (ParsedNode | string)[] = []): ParsedNode {
  return { tag, attrs, children };
}

describe('PPTX P0 semantic - style features', () => {
  it('parses gradient fill with stops and angle', () => {
    const spPr = makeNode('p:spPr', {}, [
      makeNode('a:gradFill', { type: 'linear' }, [
        makeNode('a:lin', { ang: '5400000' }),
        makeNode('a:gsLst', {}, [
          makeNode('a:gs', { pos: '0' }, [makeNode('a:srgbClr', { val: 'FF0000' })]),
          makeNode('a:gs', { pos: '100000' }, [makeNode('a:srgbClr', { val: '0000FF' })]),
        ]),
      ]),
    ]);

    const style = parseShapeStyle(spPr);
    expect(style!.fill!.type).toBe('gradient');
    expect(style!.fill!.gradientFill!.type).toBe('linear');
    expect(style!.fill!.gradientFill!.angle).toBe(90);
    expect(style!.fill!.gradientFill!.stops).toHaveLength(2);
    expect(style!.fill!.gradientFill!.stops[0]).toEqual({ position: 0, color: 'FF0000' });
    expect(style!.fill!.gradientFill!.stops[1]).toEqual({ position: 100, color: '0000FF' });
  });

  it('parses pattern fill', () => {
    const spPr = makeNode('p:spPr', {}, [
      makeNode('a:pattFill', { prst: 'cross' }, [
        makeNode('a:fgClr', {}, [makeNode('a:srgbClr', { val: 'FF0000' })]),
        makeNode('a:bgClr', {}, [makeNode('a:srgbClr', { val: 'FFFFFF' })]),
      ]),
    ]);

    const style = parseShapeStyle(spPr);
    expect(style!.fill!.type).toBe('pattern');
    expect(style!.fill!.patternFill!.preset).toBe('cross');
    expect(style!.fill!.patternFill!.fgColor).toBe('FF0000');
    expect(style!.fill!.patternFill!.bgColor).toBe('FFFFFF');
  });

  it('parses blip fill', () => {
    const spPr = makeNode('p:spPr', {}, [
      makeNode('a:blipFill', {}, [
        makeNode('a:blip', { 'r:embed': 'rId2' }),
      ]),
    ]);

    const style = parseShapeStyle(spPr);
    expect(style!.fill!.type).toBe('blip');
    expect(style!.fill!.blipRelationshipId).toBe('rId2');
  });

  it('parses noFill', () => {
    const spPr = makeNode('p:spPr', {}, [makeNode('a:noFill')]);
    const style = parseShapeStyle(spPr);
    expect(style!.fill!.type).toBe('none');
  });

  it('parses grpFill', () => {
    const spPr = makeNode('p:spPr', {}, [makeNode('a:grpFill')]);
    const style = parseShapeStyle(spPr);
    expect(style!.fill!.type).toBe('group');
  });

  it('parses line endpoints', () => {
    const spPr = makeNode('p:spPr', {}, [
      makeNode('a:ln', { w: '25400' }, [
        makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: '000000' })]),
        makeNode('a:headEnd', { type: 'triangle', w: 'med', len: 'lg' }),
        makeNode('a:tailEnd', { type: 'stealth', w: 'sm', len: 'sm' }),
      ]),
    ]);

    const style = parseShapeStyle(spPr);
    expect(style!.border!.headEnd!.type).toBe('triangle');
    expect(style!.border!.headEnd!.width).toBe('med');
    expect(style!.border!.headEnd!.length).toBe('lg');
    expect(style!.border!.tailEnd!.type).toBe('stealth');
  });

  it('parses compound line', () => {
    const spPr = makeNode('p:spPr', {}, [
      makeNode('a:ln', { cmpd: 'dbl', cap: 'round' }),
    ]);

    const style = parseShapeStyle(spPr);
    expect(style!.border!.compound).toBe('dbl');
    expect(style!.border!.cap).toBe('round');
  });

  it('parses glow effect', () => {
    const spPr = makeNode('p:spPr', {}, [
      makeNode('a:effectLst', {}, [
        makeNode('a:glow', { rad: '127000' }, [
          makeNode('a:srgbClr', { val: 'FF6600' }),
        ]),
      ]),
    ]);

    const style = parseShapeStyle(spPr);
    expect(style!.glow!.radius).toBe(10);
    expect(style!.glow!.color).toBe('FF6600');
  });

  it('parses soft edge', () => {
    const spPr = makeNode('p:spPr', {}, [
      makeNode('a:effectLst', {}, [
        makeNode('a:softEdge', { rad: '63500' }),
      ]),
    ]);

    const style = parseShapeStyle(spPr);
    expect(style!.softEdge!.radius).toBe(5);
  });

  it('parses scheme color', () => {
    const parent = makeNode('a:solidFill', {}, [
      makeNode('a:schemeClr', { val: 'accent1' }),
    ]);
    const color = parseColorValue(parent);
    expect(color).toBe('scheme:accent1');
  });

  it('parses preset color', () => {
    const parent = makeNode('a:solidFill', {}, [
      makeNode('a:prstClr', { val: 'red' }),
    ]);
    const color = parseColorValue(parent);
    expect(color).toBe('red');
  });
});
