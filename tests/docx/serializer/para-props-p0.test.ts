import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/docx/serializer.js';
import type { DocxDocument } from '../../../src/docx/types.js';

function makeDoc(props: any): DocxDocument {
  return {
    meta: {},
    styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
    body: { blocks: [{ type: 'paragraph', properties: props, runs: [{ text: 'x' }] }] },
  };
}

describe('DOCX serializer - paragraph properties P0', () => {
  it('serializes outlineLevel', () => {
    const xml = semanticToXml(makeDoc({ outlineLevel: 1 }));
    expect(xml).toContain('<w:outlineLvl');
    expect(xml).toContain('w:val="1"');
  });

  it('serializes keepNext', () => {
    const xml = semanticToXml(makeDoc({ keepNext: true }));
    expect(xml).toContain('<w:keepNext');
  });

  it('serializes keepLines', () => {
    const xml = semanticToXml(makeDoc({ keepLines: true }));
    expect(xml).toContain('<w:keepLines');
  });

  it('serializes pageBreakBefore', () => {
    const xml = semanticToXml(makeDoc({ pageBreakBefore: true }));
    expect(xml).toContain('<w:pageBreakBefore');
  });

  it('serializes paragraph shading', () => {
    const xml = semanticToXml(makeDoc({ shading: { fill: 'FFFF00', pattern: 'clear' } }));
    expect(xml).toContain('<w:shd');
    expect(xml).toContain('w:fill="FFFF00"');
  });

  it('serializes paragraph border top', () => {
    const xml = semanticToXml(makeDoc({
      border: { top: { style: 'single', size: 4, color: '000000' } }
    }));
    expect(xml).toContain('<w:pBdr');
    expect(xml).toContain('<w:top');
    expect(xml).toContain('w:val="single"');
  });

  it('serializes tab stops', () => {
    const xml = semanticToXml(makeDoc({
      tabs: [{ position: 720, alignment: 'left', leader: 'none' }]
    }));
    expect(xml).toContain('<w:tabs');
    expect(xml).toContain('<w:tab');
    expect(xml).toContain('w:pos="720"');
    expect(xml).toContain('w:val="left"');
  });
});
