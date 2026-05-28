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

describe('PPTX serializer - shape P0 features', () => {
  it('serializes preset geometry', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Rect',
        position: { x: 100, y: 100, width: 200, height: 100 },
        paragraphs: [{ type: 'paragraph', runs: [{ text: 'Rect' }] }],
        presetGeom: 'roundRect',
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('a:prstGeom');
    expect(result.slides[0]).toContain('prst="roundRect"');
  });

  it('serializes body properties', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Test',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{ type: 'paragraph', runs: [{ text: 'Test' }] }],
        bodyProperties: {
          anchor: 'ctr', wrap: 'square',
          leftInset: 0.1, topInset: 0.05, rightInset: 0.1, bottomInset: 0.05,
          vertical: 'horz', autoFit: 'normal',
        },
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('anchor="ctr"');
    expect(result.slides[0]).toContain('wrap="square"');
    expect(result.slides[0]).toContain('lIns="91440"');
    expect(result.slides[0]).toContain('tIns="45720"');
    expect(result.slides[0]).toContain('a:normAutofit');
  });

  it('serializes paragraph alignment and indentation', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Centered',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{
          type: 'paragraph',
          runs: [{ text: 'Centered' }],
          properties: {
            alignment: 'center',
            indent: { left: 18, right: 9, firstLine: -18 },
          },
        }],
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('algn="ctr"');
    expect(result.slides[0]).toContain('marL="228600"');
    expect(result.slides[0]).toContain('marR="114300"');
    expect(result.slides[0]).toContain('indent="-228600"');
  });

  it('serializes paragraph spacing', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Spaced',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{
          type: 'paragraph',
          runs: [{ text: 'Spaced' }],
          properties: {
            spacing: { before: 12, after: 6, line: 150, lineRule: 'auto' },
          },
        }],
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('a:spcBef');
    expect(result.slides[0]).toContain('val="1200"');
    expect(result.slides[0]).toContain('a:spcAft');
    expect(result.slides[0]).toContain('val="600"');
    expect(result.slides[0]).toContain('a:spcPct');
    expect(result.slides[0]).toContain('val="150000"');
  });

  it('serializes bullet character', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Bullet',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{
          type: 'paragraph',
          runs: [{ text: 'Bullet' }],
          numbering: { level: 0, numId: 'buChar', format: '•' },
        }],
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('a:buChar');
    expect(result.slides[0]).toContain('char="•"');
  });

  it('serializes run properties', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Rich',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{
          type: 'paragraph',
          runs: [{
            text: 'Rich',
            bold: true,
            italic: true,
            underline: true,
            strike: true,
            fontSize: 24,
            color: 'FF0000',
            fontFamily: 'Arial',
            superscript: true,
          }],
        }],
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('b="1"');
    expect(result.slides[0]).toContain('i="1"');
    expect(result.slides[0]).toContain('u="sng"');
    expect(result.slides[0]).toContain('strike="sngStrike"');
    expect(result.slides[0]).toContain('sz="2400"');
    expect(result.slides[0]).toContain('baseline="30000"');
    expect(result.slides[0]).toContain('val="FF0000"');
    expect(result.slides[0]).toContain('typeface="Arial"');
  });

  it('serializes position flip and rotation', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Flip',
        position: { x: 100, y: 100, width: 200, height: 100, flipH: true, flipV: true, rotation: 90 },
        paragraphs: [{ type: 'paragraph', runs: [{ text: 'Flip' }] }],
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('flipH="1"');
    expect(result.slides[0]).toContain('flipV="1"');
    expect(result.slides[0]).toContain('rot="5400000"');
  });

  it('serializes gradient fill', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Grad',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{ type: 'paragraph', runs: [{ text: 'Grad' }] }],
        style: {
          fill: {
            type: 'gradient',
            gradientFill: {
              type: 'linear',
              angle: 90,
              stops: [
                { position: 0, color: 'FF0000' },
                { position: 100, color: '0000FF' },
              ],
            },
          },
        },
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('a:gradFill');
    expect(result.slides[0]).toContain('a:gsLst');
    expect(result.slides[0]).toContain('pos="0"');
    expect(result.slides[0]).toContain('pos="100000"');
    expect(result.slides[0]).toContain('ang="5400000"');
  });

  it('serializes glow and softEdge', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Glow',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{ type: 'paragraph', runs: [{ text: 'Glow' }] }],
        style: {
          glow: { radius: 10, color: 'FF6600' },
          softEdge: { radius: 5 },
        },
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('a:glow');
    expect(result.slides[0]).toContain('rad="127000"');
    expect(result.slides[0]).toContain('a:softEdge');
    expect(result.slides[0]).toContain('rad="63500"');
  });

  it('serializes line endpoints and compound', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'text',
        content: 'Line',
        position: { x: 0, y: 0, width: 200, height: 100 },
        paragraphs: [{ type: 'paragraph', runs: [{ text: 'Line' }] }],
        style: {
          border: {
            width: 2,
            color: '000000',
            compound: 'dbl',
            cap: 'round',
            headEnd: { type: 'triangle', width: 'med', length: 'lg' },
          },
        },
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('cmpd="dbl"');
    expect(result.slides[0]).toContain('cap="round"');
    expect(result.slides[0]).toContain('a:headEnd');
    expect(result.slides[0]).toContain('type="triangle"');
    expect(result.slides[0]).toContain('w="med"');
    expect(result.slides[0]).toContain('len="lg"');
  });

  it('serializes slide background', () => {
    const pres = makePresentation([{
      background: { type: 'solid', color: 'FF0000' },
      elements: [],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:bg');
    expect(result.slides[0]).toContain('p:bgPr');
    expect(result.slides[0]).toContain('val="FF0000"');
  });

  it('serializes showMasterSp and clrMap', () => {
    const pres = makePresentation([{
      showMasterSp: false,
      showMasterPhAnim: false,
      clrMap: { bg1: 'lt2', tx1: 'dk2' },
      elements: [],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('showMasterSp="0"');
    expect(result.slides[0]).toContain('showMasterPhAnim="0"');
    expect(result.slides[0]).toContain('a:overrideClrMapping');
  });

  it('serializes group shape with children', () => {
    const pres = makePresentation([{
      elements: [{
        type: 'group',
        position: { x: 100, y: 200, width: 300, height: 400 },
        childOffset: { x: 50, y: 100 },
        childExtent: { width: 250, height: 350 },
        children: [{
          type: 'text',
          content: 'Inner',
          position: { x: 0, y: 0, width: 100, height: 50 },
          paragraphs: [{ type: 'paragraph', runs: [{ text: 'Inner' }] }],
        }],
      }],
    }]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:grpSp');
    expect(result.slides[0]).toContain('a:chOff');
    expect(result.slides[0]).toContain('a:chExt');
    expect(result.slides[0]).toContain('Inner');
  });
});
