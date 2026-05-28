import { describe, it, expect } from 'vitest';
import { parseShape } from '../../../src/pptx/parsers/shape.js';
import type { ParsedNode } from '../../../src/core/types.js';

function makeNode(tag: string, attrs: Record<string, string> = {}, children: (ParsedNode | string)[] = []): ParsedNode {
  return { tag, attrs, children };
}

describe('PPTX P0 semantic - shape features', () => {
  it('parses preset geometry', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:prstGeom', { prst: 'roundRect' }, [
          makeNode('a:avLst'),
        ]),
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Test'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    expect(shape.presetGeom).toBe('roundRect');
  });

  it('parses body properties', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr', {
          anchor: 'ctr', wrap: 'square',
          lIns: '91440', tIns: '45720', rIns: '91440', bIns: '45720',
          vert: 'vert',
        }, [
          makeNode('a:normAutofit'),
        ]),
        makeNode('a:p', {}, [
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Test'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    expect(shape.bodyProperties.anchor).toBe('ctr');
    expect(shape.bodyProperties.wrap).toBe('square');
    expect(shape.bodyProperties.leftInset).toBeCloseTo(0.1, 3);
    expect(shape.bodyProperties.topInset).toBeCloseTo(0.05, 3);
    expect(shape.bodyProperties.vertical).toBe('vert');
    expect(shape.bodyProperties.autoFit).toBe('normal');
  });

  it('parses list style', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:lstStyle', {}, [
          makeNode('a:lvl1pPr', { algn: 'l', marL: '228600' }, [
            makeNode('a:defRPr', { sz: '1800' }),
          ]),
          makeNode('a:lvl2pPr', { algn: 'l', marL: '457200' }),
        ]),
        makeNode('a:p', {}, [
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Item'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    expect(shape.listStyle.defaultParagraphProperties).toHaveLength(2);
    expect(shape.listStyle.defaultParagraphProperties[0].level).toBe(1);
    expect(shape.listStyle.defaultParagraphProperties[0].alignment).toBe('l');
    expect(shape.listStyle.defaultParagraphProperties[0].fontScale).toBe(18);
    expect(shape.listStyle.defaultParagraphProperties[1].level).toBe(2);
  });

  it('parses paragraph alignment', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [
          makeNode('a:pPr', { algn: 'ctr', marL: '228600', indent: '-228600' }),
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Centered'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    const para = shape.paragraphs[0];
    expect(para.properties.alignment).toBe('center');
    expect(para.properties.indent.left).toBe(18);
    expect(para.properties.indent.firstLine).toBe(-18);
  });

  it('parses paragraph spacing', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [
          makeNode('a:pPr', {}, [
            makeNode('a:spcBef', {}, [makeNode('a:spcPts', { val: '1200' })]),
            makeNode('a:spcAft', {}, [makeNode('a:spcPts', { val: '600' })]),
            makeNode('a:lnSpc', {}, [makeNode('a:spcPct', { val: '150000' })]),
          ]),
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Spaced'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    const props = shape.paragraphs[0].properties;
    expect(props.spacing.before).toBe(12);
    expect(props.spacing.after).toBe(6);
    expect(props.spacing.line).toBe(150);
    expect(props.spacing.lineRule).toBe('auto');
  });

  it('parses bullet character', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [
          makeNode('a:pPr', { lvl: '0' }, [
            makeNode('a:buChar', { char: '•' }),
          ]),
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Bullet'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    const numbering = shape.paragraphs[0].numbering;
    expect(numbering.level).toBe(0);
    expect(numbering.numId).toBe('buChar');
    expect(numbering.format).toBe('•');
  });

  it('parses auto numbering', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [
          makeNode('a:pPr', {}, [
            makeNode('a:buAutoNum', { type: 'arabicPeriod', startAt: '1' }),
          ]),
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Numbered'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    const numbering = shape.paragraphs[0].numbering;
    expect(numbering.numId).toBe('buAutoNum');
    expect(numbering.format).toBe('arabicPeriod');
    expect(numbering.text).toBe('1');
  });

  it('parses run properties', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [
          makeNode('a:r', {}, [
            makeNode('a:rPr', {
              b: '1', i: '1', u: 'sng', strike: 'sngStrike',
              sz: '2400', cap: 'all', spc: '100', baseline: '30000',
            }, [
              makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: 'FF0000' })]),
              makeNode('a:latin', { typeface: 'Arial' }),
            ]),
            makeNode('a:t', {}, ['Rich']),
          ]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    const run = shape.paragraphs[0].runs[0];
    expect(run.bold).toBe(true);
    expect(run.italic).toBe(true);
    expect(run.underline).toBe(true);
    expect(run.strike).toBe(true);
    expect(run.fontSize).toBe(24);
    expect(run.caps).toBe(true);
    expect(run.characterSpacing).toBe(1);
    expect(run.superscript).toBe(true);
    expect(run.color).toBe('FF0000');
    expect(run.fontFamily).toBe('Arial');
  });

  it('parses position flipH/flipV/rotation', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', { flipH: '1', flipV: '1', rot: '5400000' }, [
          makeNode('a:off', { x: '100', y: '200' }),
          makeNode('a:ext', { cx: '300', cy: '400' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [makeNode('a:r', {}, [makeNode('a:t', {}, ['Flipped'])])]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    expect(shape.position.flipH).toBe(true);
    expect(shape.position.flipV).toBe(true);
    expect(shape.position.rotation).toBe(90);
    expect(shape.rotation).toBe(90);
  });

  it('parses line break in paragraph', () => {
    const node = makeNode('p:sp', {}, [
      makeNode('p:spPr', {}, [
        makeNode('a:xfrm', {}, [
          makeNode('a:off', { x: '0', y: '0' }),
          makeNode('a:ext', { cx: '100', cy: '100' }),
        ]),
      ]),
      makeNode('p:txBody', {}, [
        makeNode('a:bodyPr'),
        makeNode('a:p', {}, [
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Line1'])]),
          makeNode('a:br'),
          makeNode('a:r', {}, [makeNode('a:t', {}, ['Line2'])]),
        ]),
      ]),
    ]);

    const shape = parseShape(node) as any;
    expect(shape.paragraphs[0].runs).toHaveLength(3);
    expect(shape.paragraphs[0].runs[1].text).toBe('\n');
  });
});
