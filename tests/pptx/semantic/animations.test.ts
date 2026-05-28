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

describe('PPTX semantic - animations', () => {
  it('parses slide transition', () => {
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
              children: [],
            },
          ],
        },
        {
          tag: 'p:transition',
          attrs: { spd: 'fast' },
          children: [
            { tag: 'p:fade', attrs: {}, children: [] },
          ],
        },
      ],
    };

    const raw = makePresentationWithSlide(slideXml);
    const semantic = rawToSemantic(raw);
    expect(semantic.slides[0].transition).toBeDefined();
    expect(semantic.slides[0].transition!.type).toBe('fade');
    expect(semantic.slides[0].transition!.duration).toBe(500);
  });

  it('parses animations from slide', () => {
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
                      tag: 'p:nvSpPr',
                      attrs: {},
                      children: [
                        { tag: 'p:cNvPr', attrs: { id: '1', name: 'Shape1' }, children: [] },
                        { tag: 'p:cNvSpPr', attrs: {}, children: [] },
                        { tag: 'p:nvPr', attrs: {}, children: [] },
                      ],
                    },
                    {
                      tag: 'p:spPr',
                      attrs: {},
                      children: [],
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
                                { tag: 'a:t', attrs: {}, children: ['Hello'] },
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
        {
          tag: 'p:timing',
          attrs: {},
          children: [
            {
              tag: 'p:tnLst',
              attrs: {},
              children: [
                {
                  tag: 'p:par',
                  attrs: {},
                  children: [
                    {
                      tag: 'p:cTn',
                      attrs: { id: '1', presetID: '1', presetClass: 'entr', triggerType: 'onClick' },
                      children: [
                        {
                          tag: 'p:childTnLst',
                          attrs: {},
                          children: [
                            {
                              tag: 'p:set',
                              attrs: {},
                              children: [
                                {
                                  tag: 'p:cBhvr',
                                  attrs: {},
                                  children: [
                                    {
                                      tag: 'p:cTn',
                                      attrs: { id: '2', dur: '1000' },
                                      children: [],
                                    },
                                    {
                                      tag: 'p:tgtEl',
                                      attrs: {},
                                      children: [
                                        { tag: 'p:spTgt', attrs: { spid: '1' }, children: [] },
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
    expect(semantic.slides).toHaveLength(1);
    expect(semantic.slides[0].animations).toHaveLength(1);
    expect(semantic.slides[0].animations![0].trigger).toBe('onClick');
    expect(semantic.slides[0].animations![0].type).toBe('parallel');
    expect(semantic.slides[0].animations![0].children).toHaveLength(1);
    expect(semantic.slides[0].animations![0].children![0].type).toBe('set');
    expect(semantic.slides[0].animations![0].children![0].shapeId).toBe('1');
    expect(semantic.slides[0].animations![0].children![0].duration).toBe(1000);
  });
});
