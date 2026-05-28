import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/docx/serializer.js';
import type { DocxDocument } from '../../../src/docx/types.js';

function makeDoc(sectionProps: any): DocxDocument {
  return {
    meta: {},
    styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
    body: { blocks: [], sectionProperties: sectionProps },
  };
}

describe('DOCX serializer - section properties P0', () => {
  it('serializes landscape orientation', () => {
    const xml = semanticToXml(makeDoc({ orientation: 'landscape' }));
    expect(xml).toContain('<w:pgSz');
    expect(xml).toContain('w:orient="landscape"');
  });

  it('serializes page number format', () => {
    const xml = semanticToXml(makeDoc({ pageNumberFormat: 'decimal', pageNumberStart: 1 }));
    expect(xml).toContain('<w:pgNumType');
    expect(xml).toContain('w:fmt="decimal"');
    expect(xml).toContain('w:start="1"');
  });

  it('serializes titlePage', () => {
    const xml = semanticToXml(makeDoc({ titlePage: true }));
    expect(xml).toContain('<w:titlePg');
  });

  it('serializes evenAndOddHeaders', () => {
    const xml = semanticToXml(makeDoc({ evenAndOddHeaders: true }));
    expect(xml).toContain('<w:evenAndOddHeaders');
  });

  it('serializes vertical alignment', () => {
    const xml = semanticToXml(makeDoc({ verticalAlign: 'center' }));
    expect(xml).toContain('<w:vAlign');
    expect(xml).toContain('w:val="center"');
  });
});
