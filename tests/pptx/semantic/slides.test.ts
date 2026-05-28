import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/pptx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

describe('PPTX semantic - slides', () => {
  it('parses presentation with slides', () => {
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
      ],
    };

    const raw: RawDocument = {
      entries: [],
      rels,
      contentTypes: [],
      parts: new Map([
        ['ppt/presentation.xml', presentationXml],
        ['ppt/slides/slide1.xml', slideXml],
      ]),
    };

    const semantic = rawToSemantic(raw);
    expect(semantic.slides).toHaveLength(1);
    expect(semantic.slides[0].elements).toHaveLength(1);
    expect(semantic.slides[0].elements[0].type).toBe('text');
  });

  it('parses slide notes', () => {
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
    rels.set('ppt/slides/_rels/slide1.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide', target: '../notesSlides/notesSlide1.xml' },
    ]);

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
                              children: [{ tag: 'a:t', attrs: {}, children: ['Slide Content'] }],
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

    const notesXml: ParsedNode = {
      tag: 'p:notes',
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
                              children: [{ tag: 'a:t', attrs: {}, children: ['Speaker notes here'] }],
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

    const raw: RawDocument = {
      entries: [],
      rels,
      contentTypes: [],
      parts: new Map([
        ['ppt/presentation.xml', presentationXml],
        ['ppt/slides/slide1.xml', slideXml],
        ['ppt/notesSlides/notesSlide1.xml', notesXml],
      ]),
    };

    const semantic = rawToSemantic(raw);
    expect(semantic.slides).toHaveLength(1);
    expect(semantic.slides[0].notes).toBe('Speaker notes here');
  });
});
