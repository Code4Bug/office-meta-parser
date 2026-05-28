import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../src/docx/semantic.js';
import { semanticToXml } from '../../src/docx/serializer.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';
import type { DocxDocument } from '../../src/docx/types.js';

function makeRaw(bodyChildren: ParsedNode[]): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map([['word/document.xml', {
      tag: 'w:document', attrs: {},
      children: [{ tag: 'w:body', attrs: {}, children: bodyChildren }],
    }]]),
  };
}

describe('DOCX bookmarks', () => {
  it('parses bookmarkStart', () => {
    const sem = rawToSemantic(makeRaw([
      { tag: 'w:bookmarkStart', attrs: { 'w:id': '1', 'w:name': 'myAnchor' }, children: [] },
      { tag: 'w:p', attrs: {}, children: [] },
    ]));
    const bookmark = sem.body.blocks[0] as any;
    expect(bookmark.type).toBe('bookmarkStart');
    expect(bookmark.id).toBe('1');
    expect(bookmark.name).toBe('myAnchor');
  });

  it('parses bookmarkEnd', () => {
    const sem = rawToSemantic(makeRaw([
      { tag: 'w:bookmarkEnd', attrs: { 'w:id': '1' }, children: [] },
    ]));
    const bookmark = sem.body.blocks[0] as any;
    expect(bookmark.type).toBe('bookmarkEnd');
    expect(bookmark.id).toBe('1');
  });

  it('serializes bookmarks round-trip', () => {
    const doc: DocxDocument = {
      meta: {},
      styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
      body: {
        blocks: [
          { type: 'bookmarkStart', id: '1', name: 'intro' } as any,
          { type: 'paragraph', runs: [{ text: 'Hello' }] },
          { type: 'bookmarkEnd', id: '1' } as any,
        ],
      },
    };
    const xml = semanticToXml(doc);
    expect(xml).toContain('w:bookmarkStart');
    expect(xml).toContain('w:name="intro"');
    expect(xml).toContain('w:bookmarkEnd');
  });
});
