import { describe, it, expect } from 'vitest';
import { unzip, zip } from '../../src/core/zip.js';

describe('unzip', () => {
  it('unpacks a ZIP buffer into entries', async () => {
    const JSZip = (await import('jszip')).default;
    const archive = new JSZip();
    archive.file('test.txt', 'hello');
    archive.file('sub/nested.txt', 'world');
    const buffer = await archive.generateAsync({ type: 'arraybuffer' });

    const entries = await unzip(buffer);
    expect(entries).toHaveLength(2);
    // Sorted alphabetically by path
    expect(entries[0].path).toBe('sub/nested.txt');
    expect(entries[1].path).toBe('test.txt');

    const decoder = new TextDecoder();
    expect(decoder.decode(entries[0].data)).toBe('world');
    expect(decoder.decode(entries[1].data)).toBe('hello');
  });
});

describe('zip', () => {
  it('packs entries into a ZIP buffer', async () => {
    const encoder = new TextEncoder();
    const entries = [
      { path: 'a.txt', data: encoder.encode('aaa').buffer },
      { path: 'b.txt', data: encoder.encode('bbb').buffer },
    ];

    const buffer = await zip(entries);
    expect(buffer).toBeInstanceOf(ArrayBuffer);

    const JSZip = (await import('jszip')).default;
    const archive = await JSZip.loadAsync(buffer);
    expect(archive.file('a.txt')).not.toBeNull();
    expect(archive.file('b.txt')).not.toBeNull();
    expect(await archive.file('a.txt')!.async('string')).toBe('aaa');
  });
});

describe('round-trip', () => {
  it('zip -> unzip preserves data', async () => {
    const encoder = new TextEncoder();
    const original = [
      { path: 'doc.xml', data: encoder.encode('<root/>').buffer },
      { path: 'rels/.rels', data: encoder.encode('<Relationships/>').buffer },
    ];

    const buffer = await zip(original);
    const restored = await unzip(buffer);

    expect(restored).toHaveLength(2);
    const decoder = new TextDecoder();
    expect(decoder.decode(restored[0].data)).toBe('<root/>');
    expect(decoder.decode(restored[1].data)).toBe('<Relationships/>');
  });
});
