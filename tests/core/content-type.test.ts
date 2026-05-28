import { describe, it, expect } from 'vitest';
import { parseContentTypes, serializeContentTypes } from '../../src/core/content-type.js';

const sampleContentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats.wordprocessingml.styles+xml"/>
</Types>`;

describe('parseContentTypes', () => {
  it('parses Default and Override entries', () => {
    const types = parseContentTypes(sampleContentTypesXml);
    expect(types).toHaveLength(4);
    expect(types[0]).toEqual({
      partName: '.rels',
      contentType: 'application/vnd.openxmlformats-package.relationships+xml',
    });
    expect(types[1]).toEqual({
      partName: '.xml',
      contentType: 'application/xml',
    });
    expect(types[2]).toEqual({
      partName: '/word/document.xml',
      contentType: 'application/vnd.openxmlformats.wordprocessingml.document.main+xml',
    });
  });
});

describe('serializeContentTypes', () => {
  it('serializes content types to XML', () => {
    const types = [
      { partName: '.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
    ];
    const xml = serializeContentTypes(types);
    expect(xml).toContain('Types');
    expect(xml).toContain('Default');
    expect(xml).toContain('Extension="rels"');
  });

  it('round-trips content types', () => {
    const types = parseContentTypes(sampleContentTypesXml);
    const xml = serializeContentTypes(types);
    const restored = parseContentTypes(xml);
    expect(restored).toEqual(types);
  });
});
