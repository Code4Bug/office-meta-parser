import { describe, it, expect } from 'vitest';
import { serializeCommentExts } from '../../../src/docx/serializer.js';
import type { CommentExtended } from '../../../src/docx/types.js';

describe('DOCX serializer - commentsExtended', () => {
  it('serializes comment extensions', () => {
    const exts: CommentExtended[] = [
      { paraId: '6D3DE01F', done: false },
      { paraId: '7E744B72', done: true },
    ];

    const xml = serializeCommentExts(exts);
    expect(xml).toContain('w15:commentsEx');
    expect(xml).toContain('w15:commentEx');
    expect(xml).toContain('w15:paraId="6D3DE01F"');
    expect(xml).toContain('w15:paraId="7E744B72"');
    expect(xml).toContain('w15:done="1"');
  });

  it('omits done when false', () => {
    const exts: CommentExtended[] = [{ paraId: 'ABC123' }];
    const xml = serializeCommentExts(exts);
    expect(xml).toContain('w15:paraId="ABC123"');
    expect(xml).not.toContain('w15:done');
  });
});
