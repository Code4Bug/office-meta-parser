import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/docx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

describe('DOCX semantic - inline elements', () => {
  it('parses hyperlink', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:hyperlink',
              attrs: { 'r:id': 'rId1', 'w:tooltip': 'Click here' },
              children: [
                {
                  tag: 'w:r',
                  attrs: {},
                  children: [
                    { tag: 'w:t', attrs: {}, children: ['Visit Google'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const rels = new Map<string, import('../../../src/core/types.js').Relationship[]>();
    rels.set('word/_rels/document.xml.rels', [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', target: 'https://www.google.com' },
    ]);

    const raw: RawDocument = {
      entries: [],
      rels,
      contentTypes: [],
      parts: new Map([['word/document.xml', documentXml]]),
    };

    const semantic = rawToSemantic(raw);
    expect(semantic.body.blocks).toHaveLength(1);
    const hyperlink = semantic.body.blocks[0] as any;
    expect(hyperlink.type).toBe('hyperlink');
    expect(hyperlink.relationshipId).toBe('rId1');
    expect(hyperlink.url).toBe('https://www.google.com');
    expect(hyperlink.tooltip).toBe('Click here');
    expect(hyperlink.runs[0].text).toBe('Visit Google');
  });

  it('parses comments', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Hello'] }] },
              ],
            },
          ],
        },
      ],
    };

    const commentsXml: ParsedNode = {
      tag: 'w:comments',
      attrs: {},
      children: [
        {
          tag: 'w:comment',
          attrs: { 'w:id': '1', 'w:author': 'John', 'w:date': '2023-01-01T00:00:00Z', 'w:initials': 'J' },
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Nice work!'] }] },
              ],
            },
          ],
        },
      ],
    };

    const raw: RawDocument = {
      entries: [],
      rels: new Map(),
      contentTypes: [],
      parts: new Map([
        ['word/document.xml', documentXml],
        ['word/comments.xml', commentsXml],
      ]),
    };

    const semantic = rawToSemantic(raw);
    expect(semantic.comments).toBeDefined();
    expect(semantic.comments).toHaveLength(1);
    expect(semantic.comments![0].id).toBe('1');
    expect(semantic.comments![0].author).toBe('John');
    expect(semantic.comments![0].content[0].runs[0].text).toBe('Nice work!');
  });

  it('parses inline image', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                {
                  tag: 'w:r',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:drawing',
                      attrs: {},
                      children: [
                        {
                          tag: 'wp:inline',
                          attrs: {},
                          children: [
                            { tag: 'wp:extent', attrs: { cx: '914400', cy: '685800' }, children: [] },
                            { tag: 'wp:docPr', attrs: { id: '1', name: 'Image 1', descr: 'A test image' }, children: [] },
                            {
                              tag: 'a:graphic',
                              attrs: {},
                              children: [
                                {
                                  tag: 'a:graphicData',
                                  attrs: {},
                                  children: [
                                    {
                                      tag: 'pic:pic',
                                      attrs: {},
                                      children: [
                                        {
                                          tag: 'pic:blipFill',
                                          attrs: {},
                                          children: [
                                            { tag: 'a:blip', attrs: { 'r:embed': 'rId1' }, children: [] },
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

    const raw: RawDocument = {
      entries: [],
      rels: new Map(),
      contentTypes: [],
      parts: new Map([['word/document.xml', documentXml]]),
    };

    const semantic = rawToSemantic(raw);
    expect(semantic.body.blocks).toHaveLength(1);
    const image = semantic.body.blocks[0] as any;
    expect(image.type).toBe('image');
    expect(image.relationshipId).toBe('rId1');
    expect(image.width).toBe(914400);
    expect(image.height).toBe(685800);
    expect(image.alt).toBe('A test image');
  });
});
