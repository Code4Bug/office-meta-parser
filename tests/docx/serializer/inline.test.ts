import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/docx/serializer.js';
import type { DocxDocument } from '../../../src/docx/types.js';

function makeDoc(blocks: any[]): DocxDocument {
  return {
    meta: {},
    styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
    body: { blocks },
  };
}

describe('DOCX serializer - inline elements', () => {
  it('serializes image', () => {
    const doc = makeDoc([
      {
        type: 'image',
        relationshipId: 'rId1',
        width: 2,
        height: 1.5,
        alt: 'Test image',
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:drawing');
    expect(xml).toContain('wp:inline');
    expect(xml).toContain('a:graphic');
    expect(xml).toContain('pic:pic');
    expect(xml).toContain('r:embed="rId1"');
  });

  it('serializes hyperlink', () => {
    const doc = makeDoc([
      {
        type: 'hyperlink',
        relationshipId: 'rId1',
        url: 'https://example.com',
        tooltip: 'Visit site',
        runs: [{ text: 'Click me' }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:hyperlink');
    expect(xml).toContain('r:id="rId1"');
    expect(xml).toContain('w:tooltip="Visit site"');
    expect(xml).toContain('Click me');
  });
});
