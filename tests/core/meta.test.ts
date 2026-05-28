import { describe, it, expect } from 'vitest';
import { parseMeta, serializeMeta, createMetaOps } from '../../src/core/meta.js';
import { parseXml } from '../../src/core/xml.js';

describe('serializeMeta', () => {
  it('serializes empty meta', () => {
    const xml = serializeMeta({});
    expect(xml).toContain('cp:coreProperties');
    expect(xml).toContain('xmlns:cp');
    expect(xml).toContain('xmlns:dc');
  });

  it('serializes title', () => {
    const xml = serializeMeta({ title: 'My Document' });
    expect(xml).toContain('dc:title');
    expect(xml).toContain('My Document');
  });

  it('serializes subject', () => {
    const xml = serializeMeta({ subject: 'Test Subject' });
    expect(xml).toContain('dc:subject');
    expect(xml).toContain('Test Subject');
  });

  it('serializes creator', () => {
    const xml = serializeMeta({ creator: 'John Doe' });
    expect(xml).toContain('dc:creator');
    expect(xml).toContain('John Doe');
  });

  it('serializes description', () => {
    const xml = serializeMeta({ description: 'A test document' });
    expect(xml).toContain('dc:description');
    expect(xml).toContain('A test document');
  });

  it('serializes keywords', () => {
    const xml = serializeMeta({ keywords: 'test, document' });
    expect(xml).toContain('cp:keywords');
    expect(xml).toContain('test, document');
  });

  it('serializes lastModifiedBy', () => {
    const xml = serializeMeta({ lastModifiedBy: 'Jane' });
    expect(xml).toContain('cp:lastModifiedBy');
    expect(xml).toContain('Jane');
  });

  it('serializes created date', () => {
    const xml = serializeMeta({ created: '2024-01-01T00:00:00Z' });
    expect(xml).toContain('dcterms:created');
    expect(xml).toContain('xsi:type="dcterms:W3CDTF"');
    expect(xml).toContain('2024-01-01T00:00:00Z');
  });

  it('serializes modified date', () => {
    const xml = serializeMeta({ modified: '2024-12-31T23:59:59Z' });
    expect(xml).toContain('dcterms:modified');
    expect(xml).toContain('xsi:type="dcterms:W3CDTF"');
    expect(xml).toContain('2024-12-31T23:59:59Z');
  });

  it('serializes revision', () => {
    const xml = serializeMeta({ revision: '5' });
    expect(xml).toContain('cp:revision');
    expect(xml).toContain('5');
  });

  it('serializes category', () => {
    const xml = serializeMeta({ category: 'Report' });
    expect(xml).toContain('cp:category');
    expect(xml).toContain('Report');
  });

  it('serializes complete meta', () => {
    const xml = serializeMeta({
      title: 'Test',
      creator: 'John',
      created: '2024-01-01T00:00:00Z',
      modified: '2024-12-31T23:59:59Z',
      revision: '3',
    });

    expect(xml).toContain('dc:title');
    expect(xml).toContain('dc:creator');
    expect(xml).toContain('dcterms:created');
    expect(xml).toContain('dcterms:modified');
    expect(xml).toContain('cp:revision');
  });
});

describe('parseMeta', () => {
  it('parses all fields', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>My Document</dc:title>
  <dc:subject>Test Subject</dc:subject>
  <dc:creator>John Doe</dc:creator>
  <dc:description>A test document</dc:description>
  <cp:keywords>test, document</cp:keywords>
  <cp:lastModifiedBy>Jane</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-12-31T23:59:59Z</dcterms:modified>
  <cp:revision>5</cp:revision>
  <cp:category>Report</cp:category>
</cp:coreProperties>`;
    const node = parseXml(xml);
    const meta = parseMeta(node);

    expect(meta.title).toBe('My Document');
    expect(meta.subject).toBe('Test Subject');
    expect(meta.creator).toBe('John Doe');
    expect(meta.description).toBe('A test document');
    expect(meta.keywords).toBe('test, document');
    expect(meta.lastModifiedBy).toBe('Jane');
    expect(meta.created).toBe('2024-01-01T00:00:00Z');
    expect(meta.modified).toBe('2024-12-31T23:59:59Z');
    expect(meta.revision).toBe('5');
    expect(meta.category).toBe('Report');
  });

  it('parses partial fields', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Only Title</dc:title>
</cp:coreProperties>`;
    const node = parseXml(xml);
    const meta = parseMeta(node);

    expect(meta.title).toBe('Only Title');
    expect(meta.creator).toBeUndefined();
    expect(meta.subject).toBeUndefined();
  });

  it('parses empty coreProperties', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`;
    const node = parseXml(xml);
    const meta = parseMeta(node);

    expect(meta.title).toBeUndefined();
    expect(meta.creator).toBeUndefined();
  });
});

describe('round-trip', () => {
  it('parse then serialize preserves all fields', () => {
    const original = {
      title: 'Test Title',
      subject: 'Test Subject',
      creator: 'John Doe',
      description: 'A description',
      keywords: 'test, meta',
      lastModifiedBy: 'Jane',
      created: '2024-01-01T00:00:00Z',
      modified: '2024-12-31T23:59:59Z',
      revision: '42',
      category: 'Report',
    };

    const xml = serializeMeta(original);
    const node = parseXml(xml);
    const parsed = parseMeta(node);

    expect(parsed.title).toBe(original.title);
    expect(parsed.subject).toBe(original.subject);
    expect(parsed.creator).toBe(original.creator);
    expect(parsed.description).toBe(original.description);
    expect(parsed.keywords).toBe(original.keywords);
    expect(parsed.lastModifiedBy).toBe(original.lastModifiedBy);
    expect(parsed.created).toBe(original.created);
    expect(parsed.modified).toBe(original.modified);
    expect(parsed.revision).toBe(original.revision);
    expect(parsed.category).toBe(original.category);
  });
});

describe('createMetaOps', () => {
  it('generates all 10 operations', () => {
    const ops = createMetaOps();
    expect(typeof ops.updateTitle).toBe('function');
    expect(typeof ops.updateSubject).toBe('function');
    expect(typeof ops.updateCreator).toBe('function');
    expect(typeof ops.updateDescription).toBe('function');
    expect(typeof ops.updateKeywords).toBe('function');
    expect(typeof ops.updateLastModifiedBy).toBe('function');
    expect(typeof ops.updateCategory).toBe('function');
    expect(typeof ops.toJSON).toBe('function');
    expect(typeof ops.toJSONString).toBe('function');
    expect(typeof ops.saveJSON).toBe('function');
  });

  it('updateTitle works on generic holder', () => {
    const ops = createMetaOps();
    const doc = { meta: { title: 'old' } };
    ops.updateTitle(doc, 'new');
    expect(doc.meta.title).toBe('new');
    expect(doc.meta.modified).toBeDefined();
  });

  it('updateCreator works on generic holder', () => {
    const ops = createMetaOps();
    const doc = { meta: {} };
    ops.updateCreator(doc, 'Alice');
    expect(doc.meta.creator).toBe('Alice');
  });

  it('toJSON strips rawXmlParts', () => {
    const ops = createMetaOps();
    const doc = { meta: { title: 'test' }, rawXmlParts: new Map([['a', 'b']]) };
    const json = ops.toJSON(doc);
    expect(json.meta.title).toBe('test');
    expect(json.rawXmlParts).toBeUndefined();
  });

  it('toJSONString produces valid JSON', () => {
    const ops = createMetaOps();
    const doc = { meta: { title: 'hello' } };
    const str = ops.toJSONString(doc, 2);
    const parsed = JSON.parse(str);
    expect(parsed.meta.title).toBe('hello');
  });
});
