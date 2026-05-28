import { describe, it, expect } from 'vitest';
import { serializeTheme } from '../../../src/docx/serializer.js';
import { parseXml } from '../../../src/core/xml.js';
import type { ThemeDefinition } from '../../../src/xlsx/types.js';

describe('DOCX serializer - theme', () => {
  it('serializes theme with color scheme and fonts', () => {
    const theme: ThemeDefinition = {
      colorScheme: { dk1: '000000', lt1: 'FFFFFF', accent1: '4F81BD' },
      majorFont: 'Arial',
      minorFont: 'Calibri',
    };

    const xml = serializeTheme(theme);
    expect(xml).toContain('a:theme');
    expect(xml).toContain('a:clrScheme');
    expect(xml).toContain('val="000000"');
    expect(xml).toContain('val="FFFFFF"');
    expect(xml).toContain('val="4F81BD"');
    expect(xml).toContain('typeface="Arial"');
    expect(xml).toContain('typeface="Calibri"');
  });

  it('round-trips through parse and serialize', () => {
    const theme: ThemeDefinition = {
      colorScheme: { dk1: '111111', accent1: 'AAAAAA', accent2: 'BBBBBB' },
      majorFont: 'Cambria',
      minorFont: 'Calibri',
    };

    const xml = serializeTheme(theme);
    const parsed = parseXml(xml);

    // 重新解析验证结构
    expect(parsed.tag).toBe('a:theme');
    const themeElements = parsed.children.find(c => typeof c !== 'string' && c.tag === 'a:themeElements');
    expect(themeElements).toBeDefined();
  });

  it('serializes empty theme', () => {
    const theme: ThemeDefinition = {};
    const xml = serializeTheme(theme);
    expect(xml).toContain('a:theme');
    expect(xml).not.toContain('a:clrScheme');
    expect(xml).not.toContain('a:fontScheme');
  });
});
