import { describe, it, expect } from 'vitest';
import type { ZipEntry, ParsedNode, Relationship, ContentType, RawDocument } from '../../src/core/types.js';

describe('core types', () => {
  it('ZipEntry has path and data', () => {
    const entry: ZipEntry = { path: 'test.xml', data: new ArrayBuffer(0) };
    expect(entry.path).toBe('test.xml');
    expect(entry.data).toBeInstanceOf(ArrayBuffer);
  });

  it('ParsedNode has tag, attrs, children', () => {
    const node: ParsedNode = {
      tag: 'w:body',
      attrs: {},
      children: [],
    };
    expect(node.tag).toBe('w:body');
    expect(node.attrs).toEqual({});
    expect(node.children).toEqual([]);
  });

  it('Relationship has id, type, target', () => {
    const rel: Relationship = {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'word/document.xml',
    };
    expect(rel.id).toBe('rId1');
    expect(rel.targetMode).toBeUndefined();
  });

  it('RawDocument has entries, rels, contentTypes, parts', () => {
    const doc: RawDocument = {
      entries: [],
      rels: new Map(),
      contentTypes: [],
      parts: new Map(),
    };
    expect(doc.entries).toEqual([]);
    expect(doc.rels).toBeInstanceOf(Map);
  });
});
