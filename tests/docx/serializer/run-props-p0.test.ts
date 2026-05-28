import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/docx/serializer.js';
import type { DocxDocument } from '../../../src/docx/types.js';

function makeDoc(runs: any[]): DocxDocument {
  return {
    meta: {},
    styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
    body: { blocks: [{ type: 'paragraph', runs }] },
  };
}

describe('DOCX serializer - run properties P0', () => {
  it('serializes highlight', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', highlight: 'yellow' }]));
    expect(xml).toContain('<w:highlight');
    expect(xml).toContain('w:val="yellow"');
  });

  it('serializes run shading', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', shadingColor: 'FFFF00', shadingPattern: 'clear' }]));
    expect(xml).toContain('<w:shd');
    expect(xml).toContain('w:fill="FFFF00"');
    expect(xml).toContain('w:val="clear"');
  });

  it('serializes caps', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', caps: true }]));
    expect(xml).toContain('<w:caps');
  });

  it('serializes smallCaps', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', smallCaps: true }]));
    expect(xml).toContain('<w:smallCaps');
  });

  it('serializes dstrike', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', dstrike: true }]));
    expect(xml).toContain('<w:dstrike');
  });

  it('serializes vanish', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', vanish: true }]));
    expect(xml).toContain('<w:vanish');
  });

  it('serializes characterSpacing', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', characterSpacing: 40 }]));
    expect(xml).toContain('<w:spacing');
    expect(xml).toContain('w:val="40"');
  });

  it('serializes kern', () => {
    const xml = semanticToXml(makeDoc([{ text: 'hi', kern: 1440 }]));
    expect(xml).toContain('<w:kern');
    expect(xml).toContain('w:val="1440"');
  });
});
