import { describe, it, expect } from 'vitest';
import type { ParsedNode, RawDocument, Relationship } from '../../../src/core/types.js';
import { rawToSemantic } from '../../../src/pptx/semantic.js';

function makeNode(tag: string, attrs: Record<string, string> = {}, children: (ParsedNode | string)[] = []): ParsedNode {
  return { tag, attrs, children };
}

function makeRaw(parts: Record<string, ParsedNode>, rels?: Map<string, Relationship[]>): RawDocument {
  return {
    entries: [],
    rels: rels || new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('PPTX P0 semantic - theme features', () => {
  it('parses fmtScheme from theme', () => {
    const themeXml = makeNode('a:theme', {}, [
      makeNode('a:themeElements', {}, [
        makeNode('a:clrScheme', { name: 'Custom' }, [
          makeNode('a:dk1', {}, [makeNode('a:srgbClr', { val: '000000' })]),
          makeNode('a:lt1', {}, [makeNode('a:srgbClr', { val: 'FFFFFF' })]),
          makeNode('a:dk2', {}, [makeNode('a:srgbClr', { val: '44546A' })]),
          makeNode('a:lt2', {}, [makeNode('a:srgbClr', { val: 'E7E6E6' })]),
          makeNode('a:accent1', {}, [makeNode('a:srgbClr', { val: '4472C4' })]),
          makeNode('a:accent2', {}, [makeNode('a:srgbClr', { val: 'ED7D31' })]),
          makeNode('a:accent3', {}, [makeNode('a:srgbClr', { val: 'A5A5A5' })]),
          makeNode('a:accent4', {}, [makeNode('a:srgbClr', { val: 'FFC000' })]),
          makeNode('a:accent5', {}, [makeNode('a:srgbClr', { val: '5B9BD5' })]),
          makeNode('a:accent6', {}, [makeNode('a:srgbClr', { val: '70AD47' })]),
          makeNode('a:hlink', {}, [makeNode('a:srgbClr', { val: '0563C1' })]),
          makeNode('a:folHlink', {}, [makeNode('a:srgbClr', { val: '954F72' })]),
        ]),
        makeNode('a:fontScheme', { name: 'Custom' }, [
          makeNode('a:majorFont', {}, [makeNode('a:latin', { typeface: 'Arial' })]),
          makeNode('a:minorFont', {}, [makeNode('a:latin', { typeface: 'Calibri' })]),
        ]),
        makeNode('a:fmtScheme', { name: 'Custom' }, [
          makeNode('a:fillStyleLst', {}, [
            makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: 'FF0000' })]),
            makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: '00FF00' })]),
            makeNode('a:gradFill', {}, [
              makeNode('a:gsLst', {}, [
                makeNode('a:gs', { pos: '0' }, [makeNode('a:srgbClr', { val: '0000FF' })]),
                makeNode('a:gs', { pos: '100000' }, [makeNode('a:srgbClr', { val: 'FFFFFF' })]),
              ]),
              makeNode('a:lin', { ang: '5400000' }),
            ]),
          ]),
          makeNode('a:lnStyleLst', {}, [
            makeNode('a:ln', { w: '25400', cmpd: 'sng', cap: 'round' }, [
              makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: '000000' })]),
            ]),
          ]),
          makeNode('a:effectStyleLst', {}, [
            makeNode('a:effectStyle', {}, [
              makeNode('a:effectLst', {}, [
                makeNode('a:glow', { rad: '63500' }, [makeNode('a:srgbClr', { val: 'FF6600' })]),
              ]),
            ]),
          ]),
          makeNode('a:bgFillStyleLst', {}, [
            makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: 'F2F2F2' })]),
          ]),
        ]),
      ]),
    ]);

    const raw = makeRaw({ 'ppt/theme/theme1.xml': themeXml });
    const sem = rawToSemantic(raw);

    expect(sem.theme).toBeDefined();
    expect(sem.theme!.formatScheme).toBeDefined();
    const fmt = sem.theme!.formatScheme!;

    expect(fmt.fillStyles).toHaveLength(3);
    expect(fmt.fillStyles![0]).toEqual({ type: 'solid', color: 'FF0000' });
    expect(fmt.fillStyles![2].type).toBe('gradient');
    expect(fmt.fillStyles![2].gradientFill!.stops).toHaveLength(2);

    expect(fmt.lineStyles).toHaveLength(1);
    expect(fmt.lineStyles![0].width).toBe(2);
    expect(fmt.lineStyles![0].compound).toBe('sng');
    expect(fmt.lineStyles![0].cap).toBe('round');
    expect(fmt.lineStyles![0].color).toBe('000000');

    expect(fmt.effectStyles).toHaveLength(1);
    expect(fmt.effectStyles![0].glow!.radius).toBe(5);
    expect(fmt.effectStyles![0].glow!.color).toBe('FF6600');

    expect(fmt.bgFillStyles).toHaveLength(1);
    expect(fmt.bgFillStyles![0]).toEqual({ type: 'solid', color: 'F2F2F2' });
  });
});
