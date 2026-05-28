import { describe, it, expect } from 'vitest';
import { parseRels, serializeRels } from '../../src/core/rels.js';

const sampleRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

describe('parseRels', () => {
  it('parses relationships from XML', () => {
    const rels = parseRels(sampleRelsXml);
    expect(rels).toHaveLength(2);
    expect(rels[0]).toEqual({
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'word/document.xml',
    });
    expect(rels[1].id).toBe('rId2');
  });

  it('handles external target mode', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`;
    const rels = parseRels(xml);
    expect(rels[0].targetMode).toBe('External');
  });

  it('returns empty array for empty Relationships', () => {
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const rels = parseRels(xml);
    expect(rels).toEqual([]);
  });
});

describe('serializeRels', () => {
  it('serializes relationships to XML', () => {
    const rels = [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', target: 'word/document.xml' },
    ];
    const xml = serializeRels(rels);
    expect(xml).toContain('Id="rId1"');
    expect(xml).toContain('Target="word/document.xml"');
    expect(xml).toContain('Relationships');
  });

  it('round-trips relationships', () => {
    const rels = parseRels(sampleRelsXml);
    const xml = serializeRels(rels);
    const restored = parseRels(xml);
    expect(restored).toEqual(rels);
  });
});
