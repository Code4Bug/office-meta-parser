import { describe, it, expect } from 'vitest';
import { parseSlide, extractSlideSize, extractNotesSize } from '../../../src/pptx/parsers/slide.js';
import type { ParsedNode, RawDocument, Relationship } from '../../../src/core/types.js';

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

describe('PPTX P0 semantic - slide features', () => {
  it('parses slide background', () => {
    const node = makeNode('p:sld', {}, [
      makeNode('p:cSld', {}, [
        makeNode('p:bg', {}, [
          makeNode('p:bgPr', {}, [
            makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: 'FF0000' })]),
          ]),
        ]),
        makeNode('p:spTree', {}, [
          makeNode('p:nvGrpSpPr', {}, [
            makeNode('p:cNvPr', { id: '1', name: '' }),
            makeNode('p:cNvGrpSpPr'),
            makeNode('p:nvPr'),
          ]),
          makeNode('p:grpSpPr'),
        ]),
      ]),
    ]);

    const slide = parseSlide(node);
    expect(slide.background).toEqual({ type: 'solid', color: 'FF0000' });
  });

  it('parses showMasterSp = false', () => {
    const node = makeNode('p:sld', { showMasterSp: '0', showMasterPhAnim: '0' }, [
      makeNode('p:cSld', {}, [
        makeNode('p:spTree', {}, [
          makeNode('p:nvGrpSpPr', {}, [
            makeNode('p:cNvPr', { id: '1', name: '' }),
            makeNode('p:cNvGrpSpPr'),
            makeNode('p:nvPr'),
          ]),
          makeNode('p:grpSpPr'),
        ]),
      ]),
    ]);

    const slide = parseSlide(node);
    expect(slide.showMasterSp).toBe(false);
    expect(slide.showMasterPhAnim).toBe(false);
  });

  it('parses color map override', () => {
    const node = makeNode('p:sld', {}, [
      makeNode('p:cSld', {}, [
        makeNode('p:spTree', {}, [
          makeNode('p:nvGrpSpPr', {}, [
            makeNode('p:cNvPr', { id: '1', name: '' }),
            makeNode('p:cNvGrpSpPr'),
            makeNode('p:nvPr'),
          ]),
          makeNode('p:grpSpPr'),
        ]),
      ]),
      makeNode('p:clrMapOvr', {}, [
        makeNode('a:overrideClrMapping', { bg1: 'lt2', tx1: 'dk2' }),
      ]),
    ]);

    const slide = parseSlide(node);
    expect(slide.clrMap).toEqual({ bg1: 'lt2', tx1: 'dk2' });
  });

  it('parses transition with advClick/advTime', () => {
    const node = makeNode('p:sld', {}, [
      makeNode('p:cSld', {}, [
        makeNode('p:spTree', {}, [
          makeNode('p:nvGrpSpPr', {}, [
            makeNode('p:cNvPr', { id: '1', name: '' }),
            makeNode('p:cNvGrpSpPr'),
            makeNode('p:nvPr'),
          ]),
          makeNode('p:grpSpPr'),
        ]),
      ]),
      makeNode('p:transition', { spd: 'fast', advClick: '0', advTm: '5000' }, [
        makeNode('p:fade', { dir: 'in' }),
      ]),
    ]);

    const slide = parseSlide(node);
    expect(slide.transition).toBeDefined();
    expect(slide.transition!.type).toBe('fade');
    expect(slide.transition!.duration).toBe(500);
    expect(slide.transition!.advClick).toBe(false);
    expect(slide.transition!.advTime).toBe(5000);
    expect(slide.transition!.direction).toBe('in');
  });

  it('parses group shape with chOff/chExt', () => {
    const node = makeNode('p:sld', {}, [
      makeNode('p:cSld', {}, [
        makeNode('p:spTree', {}, [
          makeNode('p:nvGrpSpPr', {}, [
            makeNode('p:cNvPr', { id: '1', name: '' }),
            makeNode('p:cNvGrpSpPr'),
            makeNode('p:nvPr'),
          ]),
          makeNode('p:grpSpPr'),
          makeNode('p:grpSp', {}, [
            makeNode('p:grpSpPr', {}, [
              makeNode('a:xfrm', {}, [
                makeNode('a:off', { x: '100', y: '200' }),
                makeNode('a:ext', { cx: '300', cy: '400' }),
                makeNode('a:chOff', { x: '50', y: '100' }),
                makeNode('a:chExt', { cx: '250', cy: '350' }),
              ]),
            ]),
          ]),
        ]),
      ]),
    ]);

    const slide = parseSlide(node);
    const group = slide.elements[0] as any;
    expect(group.type).toBe('group');
    expect(group.childOffset).toEqual({ x: 50, y: 100 });
    expect(group.childExtent).toEqual({ width: 250, height: 350 });
  });

  it('extracts slide size from presentation.xml', () => {
    const presXml = makeNode('p:presentation', {}, [
      makeNode('p:sldSz', { cx: '12192000', cy: '6858000' }),
    ]);

    const raw = makeRaw({ 'ppt/presentation.xml': presXml });
    const size = extractSlideSize(raw);
    expect(size).toEqual({ width: 12192000, height: 6858000 });
  });

  it('extracts notes size from presentation.xml', () => {
    const presXml = makeNode('p:presentation', {}, [
      makeNode('p:notesSz', { cx: '6858000', cy: '9144000' }),
    ]);

    const raw = makeRaw({ 'ppt/presentation.xml': presXml });
    const size = extractNotesSize(raw);
    expect(size).toEqual({ width: 6858000, height: 9144000 });
  });
});
