import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/pptx/serializer.js';
import type { PptxPresentation } from '../../../src/pptx/types.js';

function makePresentation(slides: any[]): PptxPresentation {
  return {
    meta: {},
    slides,
    masters: [],
    layouts: [],
  };
}

describe('PPTX serializer - shapes', () => {
  it('serializes slide with text shape', () => {
    const pres = makePresentation([
      {
        elements: [
          {
            type: 'text',
            content: 'Hello',
            position: { x: 100, y: 200, width: 300, height: 400 },
            paragraphs: [{ type: 'paragraph', runs: [{ text: 'Hello' }] }],
          },
        ],
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:sld');
    expect(result.slides[0]).toContain('Hello');
  });

  it('serializes slide with image shape', () => {
    const pres = makePresentation([
      {
        elements: [
          {
            type: 'image',
            relationshipId: 'rId1',
            position: { x: 100, y: 200, width: 300, height: 400 },
          },
        ],
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:pic');
    expect(result.slides[0]).toContain('a:blip');
    expect(result.slides[0]).toContain('r:embed="rId1"');
    expect(result.slides[0]).toContain('a:xfrm');
  });

  it('serializes text shape with placeholder', () => {
    const pres = makePresentation([
      {
        elements: [
          {
            type: 'text',
            content: 'Title',
            position: { x: 0, y: 0, width: 9144000, height: 1143000 },
            paragraphs: [{ type: 'paragraph', runs: [{ text: 'Title' }] }],
            placeholder: { type: 'ctrTitle', index: 1 },
          },
        ],
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:ph');
    expect(result.slides[0]).toContain('type="ctrTitle"');
    expect(result.slides[0]).toContain('idx="1"');
  });

  it('serializes text shape with style', () => {
    const pres = makePresentation([
      {
        elements: [
          {
            type: 'text',
            content: 'Styled',
            position: { x: 100, y: 100, width: 200, height: 100 },
            paragraphs: [{ type: 'paragraph', runs: [{ text: 'Styled' }] }],
            style: {
              fill: { type: 'solid', color: 'FF0000' },
              border: { width: 1, color: '000000' },
              shadow: { blur: 4, offsetX: 2, offsetY: 2, color: '808080' },
            },
          },
        ],
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('a:solidFill');
    expect(result.slides[0]).toContain('a:srgbClr');
    expect(result.slides[0]).toContain('val="FF0000"');
    expect(result.slides[0]).toContain('a:ln');
    expect(result.slides[0]).toContain('w="12700"');
    expect(result.slides[0]).toContain('val="000000"');
    expect(result.slides[0]).toContain('a:effectLst');
    expect(result.slides[0]).toContain('a:outerShdw');
    expect(result.slides[0]).toContain('blurRad="50800"');
  });

  it('serializes slide with table', () => {
    const pres = makePresentation([
      {
        elements: [
          {
            type: 'table',
            position: { x: 100, y: 200, width: 800, height: 400 },
            rows: [
              {
                cells: [
                  { content: [{ type: 'paragraph', runs: [{ text: 'A1' }] }] },
                  { content: [{ type: 'paragraph', runs: [{ text: 'B1' }] }] },
                ],
              },
              {
                cells: [
                  { content: [{ type: 'paragraph', runs: [{ text: 'A2' }] }] },
                  { content: [{ type: 'paragraph', runs: [{ text: 'B2' }] }] },
                ],
              },
            ],
          },
        ],
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:graphicFrame');
    expect(result.slides[0]).toContain('a:tbl');
    expect(result.slides[0]).toContain('a:tr');
    expect(result.slides[0]).toContain('a:tc');
    expect(result.slides[0]).toContain('A1');
    expect(result.slides[0]).toContain('B2');
  });

  it('serializes text shape with hyperlink', () => {
    const pres = makePresentation([
      {
        elements: [
          {
            type: 'text',
            content: 'Click me',
            position: { x: 100, y: 100, width: 200, height: 50 },
            paragraphs: [{ type: 'paragraph', runs: [{ text: 'Click me' }] }],
            hyperlink: { url: 'rId1', tooltip: 'Visit site' },
          },
        ],
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('a:hlinkClick');
    expect(result.slides[0]).toContain('r:id="rId1"');
    expect(result.slides[0]).toContain('tooltip="Visit site"');
  });
});
