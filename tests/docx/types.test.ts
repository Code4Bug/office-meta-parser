import { describe, it, expect } from 'vitest';
import type {
  DocxDocument,
  DocumentMeta,
  Paragraph,
  TextRun,
  Table,
  Image,
  Comment,
  Revision,
} from '../../src/docx/types.js';

describe('DOCX types', () => {
  it('DocxDocument structure', () => {
    const doc: DocxDocument = {
      meta: { title: 'Test' },
      styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
      body: { blocks: [] },
    };
    expect(doc.meta.title).toBe('Test');
    expect(doc.body.blocks).toEqual([]);
  });

  it('Paragraph structure', () => {
    const p: Paragraph = {
      type: 'paragraph',
      runs: [{ text: 'Hello' }],
    };
    expect(p.type).toBe('paragraph');
    expect(p.runs[0].text).toBe('Hello');
  });

  it('TextRun with styles', () => {
    const run: TextRun = {
      text: 'Bold',
      bold: true,
      fontSize: 24,
      color: 'FF0000',
    };
    expect(run.bold).toBe(true);
    expect(run.fontSize).toBe(24);
  });

  it('Table structure', () => {
    const table: Table = {
      type: 'table',
      rows: [
        { cells: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Cell' }] }] }] },
      ],
    };
    expect(table.rows).toHaveLength(1);
  });

  it('Image structure', () => {
    const img: Image = {
      type: 'image',
      relationshipId: 'rId1',
      width: 100,
      height: 200,
    };
    expect(img.relationshipId).toBe('rId1');
  });

  it('Comment structure', () => {
    const comment: Comment = {
      id: '1',
      author: 'John',
      date: '2024-01-01T00:00:00Z',
      content: [{ type: 'paragraph', runs: [{ text: 'Nice work' }] }],
    };
    expect(comment.author).toBe('John');
  });

  it('Revision structure', () => {
    const rev: Revision = {
      type: 'insert',
      author: 'Jane',
      date: '2024-01-01T00:00:00Z',
    };
    expect(rev.type).toBe('insert');
  });
});
