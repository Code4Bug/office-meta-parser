import { describe, it, expect } from 'vitest';
import { serializeEndnotes } from '../../../src/docx/serializer.js';
import type { Footnote } from '../../../src/docx/types.js';

describe('DOCX serializer - endnotes', () => {
  it('serializes endnotes', () => {
    const endnotes: Footnote[] = [
      {
        id: '2',
        content: [{ type: 'paragraph', runs: [{ text: '这是尾注' }] }],
      },
      {
        id: '3',
        type: 'normal',
        content: [{ type: 'paragraph', runs: [{ text: '另一条尾注' }] }],
      },
    ];

    const xml = serializeEndnotes(endnotes);
    expect(xml).toContain('w:endnotes');
    expect(xml).toContain('w:endnote');
    expect(xml).toContain('w:id="2"');
    expect(xml).toContain('w:id="3"');
    expect(xml).toContain('w:type="normal"');
    expect(xml).toContain('这是尾注');
    expect(xml).toContain('另一条尾注');
  });

  it('serializes endnote without type', () => {
    const endnotes: Footnote[] = [
      { id: '5', content: [{ type: 'paragraph', runs: [{ text: 'test' }] }] },
    ];

    const xml = serializeEndnotes(endnotes);
    expect(xml).toContain('w:id="5"');
    expect(xml).not.toContain('w:type');
  });
});
