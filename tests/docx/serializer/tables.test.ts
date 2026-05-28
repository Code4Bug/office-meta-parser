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

describe('DOCX serializer - tables', () => {
  it('serializes a simple table', () => {
    const doc = makeDoc([
      {
        type: 'table',
        rows: [
          {
            cells: [
              { blocks: [{ type: 'paragraph', runs: [{ text: 'A1' }] }] },
              { blocks: [{ type: 'paragraph', runs: [{ text: 'B1' }] }] },
            ],
          },
          {
            cells: [
              { blocks: [{ type: 'paragraph', runs: [{ text: 'A2' }] }] },
              { blocks: [{ type: 'paragraph', runs: [{ text: 'B2' }] }] },
            ],
          },
        ],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:tbl');
    expect(xml).toContain('w:tr');
    expect(xml).toContain('w:tc');
    expect(xml).toContain('A1');
    expect(xml).toContain('B2');
  });

  it('serializes table with borders', () => {
    const doc = makeDoc([
      {
        type: 'table',
        properties: {
          width: 5000,
          borders: {
            top: { style: 'single', size: 4, color: '000000' },
            bottom: { style: 'single', size: 4, color: '000000' },
          },
        },
        rows: [
          {
            cells: [
              { blocks: [{ type: 'paragraph', runs: [{ text: 'Cell' }] }] },
            ],
          },
        ],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:tblW');
    expect(xml).toContain('w:tblBorders');
    expect(xml).toContain('w:top');
    expect(xml).toContain('w:bottom');
  });
});
