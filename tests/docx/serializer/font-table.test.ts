import { describe, it, expect } from 'vitest';
import { serializeFontTable } from '../../../src/docx/serializer.js';
import type { FontEntry } from '../../../src/docx/types.js';

describe('DOCX serializer - fontTable', () => {
  it('serializes fonts', () => {
    const fonts: FontEntry[] = [
      { name: 'Times New Roman', charset: '00', family: 'roman', pitch: 'variable' },
      { name: '宋体', altName: '汉仪书宋二KW', charset: '86', family: 'auto' },
    ];

    const xml = serializeFontTable(fonts);
    expect(xml).toContain('w:fonts');
    expect(xml).toContain('w:name="Times New Roman"');
    expect(xml).toContain('w:val="00"');
    expect(xml).toContain('w:val="roman"');
    expect(xml).toContain('w:name="宋体"');
    expect(xml).toContain('w:val="汉仪书宋二KW"');
  });

  it('serializes font with minimal fields', () => {
    const fonts: FontEntry[] = [{ name: 'Arial' }];
    const xml = serializeFontTable(fonts);
    expect(xml).toContain('w:name="Arial"');
    expect(xml).not.toContain('w:altName');
    expect(xml).not.toContain('w:charset');
  });
});
