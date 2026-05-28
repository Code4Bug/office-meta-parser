import { describe, it, expect } from 'vitest';
import { extractStyles } from '../../src/docx/parsers/styles.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('extractStyles', () => {
  it('returns empty when no styles.xml', () => {
    const raw = makeRaw({});
    const styles = extractStyles(raw);
    expect(styles.paragraphStyles).toEqual([]);
    expect(styles.characterStyles).toEqual([]);
    expect(styles.tableStyles).toEqual([]);
  });

  it('parses paragraph styles', () => {
    const stylesXml: ParsedNode = {
      tag: 'w:styles',
      attrs: {},
      children: [
        {
          tag: 'w:style',
          attrs: { 'w:type': 'paragraph', 'w:styleId': 'Heading1' },
          children: [
            { tag: 'w:name', attrs: { 'w:val': 'heading 1' }, children: [] },
            { tag: 'w:basedOn', attrs: { 'w:val': 'Normal' }, children: [] },
            { tag: 'w:next', attrs: { 'w:val': 'Normal' }, children: [] },
            {
              tag: 'w:pPr',
              attrs: {},
              children: [
                { tag: 'w:jc', attrs: { 'w:val': 'left' }, children: [] },
                { tag: 'w:spacing', attrs: { 'w:before': '240', 'w:after': '120' }, children: [] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.paragraphStyles).toHaveLength(1);
    expect(styles.paragraphStyles[0].id).toBe('Heading1');
    expect(styles.paragraphStyles[0].name).toBe('heading 1');
    expect(styles.paragraphStyles[0].basedOn).toBe('Normal');
    expect(styles.paragraphStyles[0].next).toBe('Normal');
    expect(styles.paragraphStyles[0].properties?.alignment).toBe('left');
    expect(styles.paragraphStyles[0].properties?.spacing?.before).toBe(240);
  });

  it('parses character styles', () => {
    const stylesXml: ParsedNode = {
      tag: 'w:styles',
      attrs: {},
      children: [
        {
          tag: 'w:style',
          attrs: { 'w:type': 'character', 'w:styleId': 'Strong' },
          children: [
            { tag: 'w:name', attrs: { 'w:val': 'Strong' }, children: [] },
            {
              tag: 'w:rPr',
              attrs: {},
              children: [
                { tag: 'w:b', attrs: {}, children: [] },
                { tag: 'w:color', attrs: { 'w:val': 'FF0000' }, children: [] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.characterStyles).toHaveLength(1);
    expect(styles.characterStyles[0].id).toBe('Strong');
    expect(styles.characterStyles[0].properties?.bold).toBe(true);
    expect(styles.characterStyles[0].properties?.color).toBe('FF0000');
  });

  it('parses table styles', () => {
    const stylesXml: ParsedNode = {
      tag: 'w:styles',
      attrs: {},
      children: [
        {
          tag: 'w:style',
          attrs: { 'w:type': 'table', 'w:styleId': 'TableGrid' },
          children: [
            { tag: 'w:name', attrs: { 'w:val': 'Table Grid' }, children: [] },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.tableStyles).toHaveLength(1);
    expect(styles.tableStyles[0].id).toBe('TableGrid');
    expect(styles.tableStyles[0].name).toBe('Table Grid');
  });

  it('parses mixed styles', () => {
    const stylesXml: ParsedNode = {
      tag: 'w:styles',
      attrs: {},
      children: [
        {
          tag: 'w:style',
          attrs: { 'w:type': 'paragraph', 'w:styleId': 'Normal' },
          children: [
            { tag: 'w:name', attrs: { 'w:val': 'Normal' }, children: [] },
          ],
        },
        {
          tag: 'w:style',
          attrs: { 'w:type': 'character', 'w:styleId': 'Emphasis' },
          children: [
            { tag: 'w:name', attrs: { 'w:val': 'Emphasis' }, children: [] },
            {
              tag: 'w:rPr',
              attrs: {},
              children: [
                { tag: 'w:i', attrs: {}, children: [] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/styles.xml': stylesXml });
    const styles = extractStyles(raw);
    expect(styles.paragraphStyles).toHaveLength(1);
    expect(styles.characterStyles).toHaveLength(1);
    expect(styles.characterStyles[0].properties?.italic).toBe(true);
  });
});
