import { describe, it, expect } from 'vitest';
import { serializeRevisionMark } from '../../../src/docx/serializer.js';
import type { Revision } from '../../../src/docx/types.js';

describe('DOCX serializer - revisions', () => {
  it('serializes insert revision', () => {
    const revision: Revision = {
      type: 'insert',
      author: 'John',
      date: '2024-01-01T00:00:00Z',
    };

    const node = serializeRevisionMark(revision);
    expect(node.tag).toBe('w:ins');
    expect(node.attrs['w:author']).toBe('John');
    expect(node.attrs['w:date']).toBe('2024-01-01T00:00:00Z');
  });

  it('serializes delete revision', () => {
    const revision: Revision = {
      type: 'delete',
      author: 'Jane',
      date: '2024-01-02T00:00:00Z',
    };

    const node = serializeRevisionMark(revision);
    expect(node.tag).toBe('w:del');
    expect(node.attrs['w:author']).toBe('Jane');
  });

  it('serializes format change revision', () => {
    const revision: Revision = {
      type: 'formatChange',
      author: 'Admin',
      date: '2024-01-03T00:00:00Z',
    };

    const node = serializeRevisionMark(revision);
    expect(node.tag).toBe('w:rPrChange');
  });
});
