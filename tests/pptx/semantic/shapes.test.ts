import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/pptx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makePresentationWithSlide(slideXml: ParsedNode): RawDocument {
  const presentationXml: ParsedNode = {
    tag: 'p:presentation',
    attrs: {},
    children: [
      {
        tag: 'p:sldIdLst',
        attrs: {},
        children: [
          { tag: 'p:sldId', attrs: { id: '256', 'r:id': 'rId2' }, children: [] },
        ],
      },
    ],
  };

  const rels = new Map<string, import('../../../src/core/types.js').Relationship[]>();
  rels.set('ppt/_rels/presentation.xml.rels', [
    { id: 'rId2', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', target: 'slides/slide1.xml' },
  ]);

  return {
    entries: [],
    rels,
    contentTypes: [],
    parts: new Map([
      ['ppt/presentation.xml', presentationXml],
      ['ppt/slides/slide1.xml', slideXml],
    ]),
  };
}

describe('PPTX semantic - shapes', () => {
  it('parses shape position', () => {
    const slideXml: ParsedNode = {
      tag: 'p:sld',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: {},
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
                            { tag: 'a:off', attrs: { x: '1000000', y: '2000000' }, children: [] },
                            { tag: 'a:ext', attrs: { cx: '3000000', cy: '4000000' }, children: [] },
                          ],
                        },
                      ],
                    },
                    {
                      tag: 'p:txBody',
                      attrs: {},
                      children: [
                        {
                          tag: 'a:p',
                          attrs: {},
                          children: [
                            {
                              tag: 'a:r',
                              attrs: {},
                              children: [
                                { tag: 'a:t', attrs: {}, children: ['Positioned'] },
                              ],
                            },
                          ],
                        },
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

    const raw = makePresentationWithSlide(slideXml);
    const semantic = rawToSemantic(raw);
    const shape = semantic.slides[0].elements[0] as any;
    expect(shape.position).toEqual({
      x: 1000000,
      y: 2000000,
      width: 3000000,
      height: 4000000,
    });
  });

  it('parses image shape', () => {
    const slideXml: ParsedNode = {
      tag: 'p:sld',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: {},
          children: [
            {
              tag: 'p:spTree',
              attrs: {},
              children: [
                {
                  tag: 'p:pic',
                  attrs: {},
                  children: [
                    {
                      tag: 'p:blipFill',
                      attrs: {},
                      children: [
                        { tag: 'a:blip', attrs: { 'r:embed': 'rId1' }, children: [] },
                      ],
                    },
                    {
                      tag: 'p:spPr',
                      attrs: {},
                      children: [
                        {
                          tag: 'a:xfrm',
                          attrs: {},
                          children: [
                            { tag: 'a:off', attrs: { x: '500000', y: '600000' }, children: [] },
                            { tag: 'a:ext', attrs: { cx: '2000000', cy: '1500000' }, children: [] },
                          ],
                        },
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

    const raw = makePresentationWithSlide(slideXml);
    const semantic = rawToSemantic(raw);
    expect(semantic.slides[0].elements).toHaveLength(1);
    const img = semantic.slides[0].elements[0] as any;
    expect(img.type).toBe('image');
    expect(img.relationshipId).toBe('rId1');
    expect(img.position).toEqual({
      x: 500000,
      y: 600000,
      width: 2000000,
      height: 1500000,
    });
  });

  it('parses placeholder', () => {
    const slideXml: ParsedNode = {
      tag: 'p:sld',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: {},
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
                        { tag: 'p:ph', attrs: { type: 'ctrTitle', idx: '1' }, children: [] },
                      ],
                    },
                    {
                      tag: 'p:txBody',
                      attrs: {},
                      children: [
                        {
                          tag: 'a:p',
                          attrs: {},
                          children: [
                            {
                              tag: 'a:r',
                              attrs: {},
                              children: [
                                { tag: 'a:t', attrs: {}, children: ['Title Placeholder'] },
                              ],
                            },
                          ],
                        },
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

    const raw = makePresentationWithSlide(slideXml);
    const semantic = rawToSemantic(raw);
    const shape = semantic.slides[0].elements[0] as any;
    expect(shape.type).toBe('text');
    expect(shape.placeholder).toBeDefined();
    expect(shape.placeholder.type).toBe('ctrTitle');
    expect(shape.placeholder.index).toBe(1);
    expect(shape.content).toBe('Title Placeholder');
  });

  it('parses group shape', () => {
    const slideXml: ParsedNode = {
      tag: 'p:sld',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: {},
          children: [
            {
              tag: 'p:spTree',
              attrs: {},
              children: [
                {
                  tag: 'p:grpSp',
                  attrs: {},
                  children: [
                    {
                      tag: 'p:grpSpPr',
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
                      ],
                    },
                    {
                      tag: 'p:sp',
                      attrs: {},
                      children: [
                        {
                          tag: 'p:txBody',
                          attrs: {},
                          children: [
                            {
                              tag: 'a:p',
                              attrs: {},
                              children: [
                                {
                                  tag: 'a:r',
                                  attrs: {},
                                  children: [{ tag: 'a:t', attrs: {}, children: ['Child'] }],
                                },
                              ],
                            },
                          ],
                        },
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

    const raw = makePresentationWithSlide(slideXml);
    const semantic = rawToSemantic(raw);
    expect(semantic.slides[0].elements).toHaveLength(1);
    const group = semantic.slides[0].elements[0] as any;
    expect(group.type).toBe('group');
    expect(group.position).toEqual({ x: 100000, y: 200000, width: 5000000, height: 3000000 });
    expect(group.children).toHaveLength(1);
    expect(group.children[0].type).toBe('text');
  });

  it('parses shape hyperlink', () => {
    const slideXml: ParsedNode = {
      tag: 'p:sld',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: {},
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
                            { tag: 'a:ext', attrs: { cx: '1000000', cy: '500000' }, children: [] },
                          ],
                        },
                        { tag: 'a:hlinkClick', attrs: { 'r:id': 'rId1', tooltip: 'Click me' }, children: [] },
                      ],
                    },
                    {
                      tag: 'p:txBody',
                      attrs: {},
                      children: [
                        {
                          tag: 'a:p',
                          attrs: {},
                          children: [
                            {
                              tag: 'a:r',
                              attrs: {},
                              children: [{ tag: 'a:t', attrs: {}, children: ['Link text'] }],
                            },
                          ],
                        },
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

    const raw = makePresentationWithSlide(slideXml);
    const semantic = rawToSemantic(raw);
    const shape = semantic.slides[0].elements[0] as any;
    expect(shape.hyperlink).toBeDefined();
    expect(shape.hyperlink.url).toBe('rId1');
    expect(shape.hyperlink.tooltip).toBe('Click me');
  });

  it('parses table shape', () => {
    const slideXml: ParsedNode = {
      tag: 'p:sld',
      attrs: {},
      children: [
        {
          tag: 'p:cSld',
          attrs: {},
          children: [
            {
              tag: 'p:spTree',
              attrs: {},
              children: [
                {
                  tag: 'p:graphicFrame',
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
                            { tag: 'a:off', attrs: { x: '500000', y: '600000' }, children: [] },
                            { tag: 'a:ext', attrs: { cx: '4000000', cy: '2000000' }, children: [] },
                          ],
                        },
                      ],
                    },
                    {
                      tag: 'a:graphic',
                      attrs: {},
                      children: [
                        {
                          tag: 'a:graphicData',
                          attrs: {},
                          children: [
                            {
                              tag: 'a:tbl',
                              attrs: {},
                              children: [
                                {
                                  tag: 'a:tr',
                                  attrs: { h: '500000' },
                                  children: [
                                    {
                                      tag: 'a:tc',
                                      attrs: { w: '2000000' },
                                      children: [
                                        {
                                          tag: 'a:txBody',
                                          attrs: {},
                                          children: [
                                            {
                                              tag: 'a:p',
                                              attrs: {},
                                              children: [
                                                {
                                                  tag: 'a:r',
                                                  attrs: {},
                                                  children: [{ tag: 'a:t', attrs: {}, children: ['Cell A1'] }],
                                                },
                                              ],
                                            },
                                          ],
                                        },
                                      ],
                                    },
                                    {
                                      tag: 'a:tc',
                                      attrs: { w: '2000000' },
                                      children: [
                                        {
                                          tag: 'a:txBody',
                                          attrs: {},
                                          children: [
                                            {
                                              tag: 'a:p',
                                              attrs: {},
                                              children: [
                                                {
                                                  tag: 'a:r',
                                                  attrs: {},
                                                  children: [{ tag: 'a:t', attrs: {}, children: ['Cell B1'] }],
                                                },
                                              ],
                                            },
                                          ],
                                        },
                                      ],
                                    },
                                  ],
                                },
                                {
                                  tag: 'a:tr',
                                  attrs: { h: '500000' },
                                  children: [
                                    {
                                      tag: 'a:tc',
                                      attrs: { w: '2000000', gridSpan: '2' },
                                      children: [
                                        {
                                          tag: 'a:txBody',
                                          attrs: {},
                                          children: [
                                            {
                                              tag: 'a:p',
                                              attrs: {},
                                              children: [
                                                {
                                                  tag: 'a:r',
                                                  attrs: {},
                                                  children: [{ tag: 'a:t', attrs: {}, children: ['Merged'] }],
                                                },
                                              ],
                                            },
                                          ],
                                        },
                                      ],
                                    },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
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

    const raw = makePresentationWithSlide(slideXml);
    const semantic = rawToSemantic(raw);
    expect(semantic.slides[0].elements).toHaveLength(1);
    const table = semantic.slides[0].elements[0] as any;
    expect(table.type).toBe('table');
    expect(table.position).toEqual({ x: 500000, y: 600000, width: 4000000, height: 2000000 });
    expect(table.rows).toHaveLength(2);
    expect(table.rows[0].cells).toHaveLength(2);
    expect(table.rows[0].height).toBe(500000);
    expect(table.rows[0].cells[0].width).toBe(2000000);
    expect(table.rows[0].cells[0].content[0].runs[0].text).toBe('Cell A1');
    expect(table.rows[0].cells[1].content[0].runs[0].text).toBe('Cell B1');
    expect(table.rows[1].cells[0].colspan).toBe(2);
    expect(table.rows[1].cells[0].content[0].runs[0].text).toBe('Merged');
  });
});
