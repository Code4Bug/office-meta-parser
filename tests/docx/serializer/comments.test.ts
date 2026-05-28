import { describe, it, expect } from 'vitest';
import { serializeComments } from '../../../src/docx/serializer.js';
import type { Comment } from '../../../src/docx/types.js';

describe('DOCX serializer - comments', () => {
  it('serializes comments', () => {
    const comments: Comment[] = [
      {
        id: '1',
        author: 'John',
        date: '2024-01-01T00:00:00Z',
        content: [{ type: 'paragraph', runs: [{ text: 'This is a comment' }] }],
        initials: 'J',
      },
    ];

    const xml = serializeComments(comments);
    expect(xml).toContain('w:comments');
    expect(xml).toContain('w:comment');
    expect(xml).toContain('w:id="1"');
    expect(xml).toContain('w:author="John"');
    expect(xml).toContain('w:initials="J"');
    expect(xml).toContain('This is a comment');
  });
});
