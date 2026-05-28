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

describe('DOCX serializer - paragraphs', () => {
  it('serializes a simple paragraph', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Hello' }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:document');
    expect(xml).toContain('w:body');
    expect(xml).toContain('w:p');
    expect(xml).toContain('w:r');
    expect(xml).toContain('w:t');
    expect(xml).toContain('Hello');
  });

  it('serializes bold text', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Bold', bold: true }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:b');
  });

  it('serializes italic text', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Italic', italic: true }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:i');
  });

  it('serializes underline text', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Underline', underline: true }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:u');
    expect(xml).toContain('val="single"');
  });

  it('serializes strikethrough text', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Strike', strike: true }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:strike');
  });

  it('serializes font size', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Large', fontSize: 24 }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:sz');
    expect(xml).toContain('w:val="24"');
  });

  it('serializes font color', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Red', color: 'FF0000' }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:color');
    expect(xml).toContain('w:val="FF0000"');
  });

  it('serializes font family', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: 'Arial', fontFamily: 'Arial' }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:rFonts');
    expect(xml).toContain('w:ascii="Arial"');
    expect(xml).toContain('w:hAnsi="Arial"');
  });

  it('serializes superscript', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: '2', superscript: true }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:vertAlign');
    expect(xml).toContain('superscript');
  });

  it('serializes subscript', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        runs: [{ text: '2', subscript: true }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:vertAlign');
    expect(xml).toContain('subscript');
  });

  it('serializes paragraph alignment', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        properties: { alignment: 'center' },
        runs: [{ text: 'Centered' }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:jc');
    expect(xml).toContain('center');
  });

  it('serializes paragraph indent', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        properties: { indent: { left: 720, firstLine: 480 } },
        runs: [{ text: 'Indented' }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:ind');
  });

  it('serializes paragraph spacing', () => {
    const doc = makeDoc([
      {
        type: 'paragraph',
        properties: { spacing: { before: 120, after: 120, line: 360, lineRule: 'auto' } },
        runs: [{ text: 'Spaced' }],
      },
    ]);

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:spacing');
  });

  it('serializes multiple paragraphs', () => {
    const doc = makeDoc([
      { type: 'paragraph', runs: [{ text: 'P1' }] },
      { type: 'paragraph', runs: [{ text: 'P2' }] },
    ]);

    const xml = semanticToXml(doc);
    const pCount = (xml.match(/<w:p>/g) || []).length;
    expect(pCount).toBe(2);
  });
});
