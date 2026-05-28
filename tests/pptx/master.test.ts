import { describe, it, expect } from 'vitest';
import { extractMasters, extractLayouts } from '../../src/pptx/parsers/master.js';
import type { RawDocument, ParsedNode, Relationship } from '../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>, rels?: Map<string, Relationship[]>): RawDocument {
  return {
    entries: [],
    rels: rels || new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('extractMasters', () => {
  it('returns empty when no slide masters', () => {
    const raw = makeRaw({});
    expect(extractMasters(raw)).toEqual([]);
  });

  it('parses slide master with background', () => {
    const masterXml: ParsedNode = {
      tag: 'p:sldMaster',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: {},
          children: [
            {
              tag: 'p:bg',
              attrs: {},
              children: [
                {
                  tag: 'p:bgPr',
                  attrs: {},
                  children: [
                    {
                      tag: 'a:solidFill',
                      attrs: {},
                      children: [
                        { tag: 'a:srgbClr', attrs: { val: 'FFFFFF' }, children: [] },
                      ],
                    },
                  ],
                },
              ],
            },
            {
              tag: 'p:spTree',
              attrs: {},
              children: [],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'ppt/slideMasters/slideMaster1.xml': masterXml });
    const masters = extractMasters(raw);
    expect(masters).toHaveLength(1);
    expect(masters[0].id).toBe('1');
    expect(masters[0].background?.type).toBe('solid');
    expect(masters[0].background?.color).toBe('FFFFFF');
  });

  it('parses slide master with associated layouts', () => {
    const masterXml: ParsedNode = {
      tag: 'p:sldMaster',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: { name: 'Title Slide' },
          children: [
            {
              tag: 'p:spTree',
              attrs: {},
              children: [],
            },
          ],
        },
      ],
    };

    const layoutXml: ParsedNode = {
      tag: 'p:sldLayout',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: { name: 'Title Layout' },
          children: [
            {
              tag: 'p:spTree',
              attrs: {},
              children: [
                {
                  tag: 'p:sp',
                  attrs: {},
                  children: [
                    {
                      tag: 'p:spPr',
                      attrs: {},
                      children: [
                        {
                          tag: 'a:xfrm',
                          attrs: {},
                          children: [
                            { tag: 'a:off', attrs: { x: '0', y: '0' }, children: [] },
                            { tag: 'a:ext', attrs: { cx: '9144000', cy: '1143000' }, children: [] },
                          ],
                        },
                        { tag: 'p:ph', attrs: { type: 'ctrTitle', idx: '0' }, children: [] },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const rels = new Map<string, Relationship[]>();
    rels.set('ppt/slideMasters/_rels/slideMaster1.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', target: '../slideLayouts/slideLayout1.xml' },
    ]);

    const raw = makeRaw({
      'ppt/slideMasters/slideMaster1.xml': masterXml,
      'ppt/slideLayouts/slideLayout1.xml': layoutXml,
    }, rels);

    const masters = extractMasters(raw);
    expect(masters).toHaveLength(1);
    expect(masters[0].layouts).toHaveLength(1);
    expect(masters[0].layouts[0].name).toBe('Title Layout');
    expect(masters[0].layouts[0].placeholders).toHaveLength(1);
    expect(masters[0].layouts[0].placeholders[0].type).toBe('ctrTitle');
  });
});

describe('extractLayouts', () => {
  it('returns empty when no slide layouts', () => {
    const raw = makeRaw({});
    expect(extractLayouts(raw)).toEqual([]);
  });

  it('parses slide layout with placeholders', () => {
    const layoutXml: ParsedNode = {
      tag: 'p:sldLayout',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: { name: 'Blank' },
          children: [
            {
              tag: 'p:spTree',
              attrs: {},
              children: [
                {
                  tag: 'p:sp',
                  attrs: {},
                  children: [
                    {
                      tag: 'p:spPr',
                      attrs: {},
                      children: [
                        {
                          tag: 'a:xfrm',
                          attrs: {},
                          children: [
                            { tag: 'a:off', attrs: { x: '100000', y: '200000' }, children: [] },
                            { tag: 'a:ext', attrs: { cx: '5000000', cy: '3000000' }, children: [] },
                          ],
                        },
                        { tag: 'p:ph', attrs: { type: 'body', idx: '1' }, children: [] },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'ppt/slideLayouts/slideLayout1.xml': layoutXml });
    const layouts = extractLayouts(raw);
    expect(layouts).toHaveLength(1);
    expect(layouts[0].id).toBe('1');
    expect(layouts[0].name).toBe('Blank');
    expect(layouts[0].placeholders).toHaveLength(1);
    expect(layouts[0].placeholders[0].type).toBe('body');
    expect(layouts[0].placeholders[0].index).toBe(1);
    expect(layouts[0].placeholders[0].position).toEqual({ x: 100000, y: 200000, width: 5000000, height: 3000000 });
  });
});
