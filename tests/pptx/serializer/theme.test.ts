import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/pptx/serializer.js';
import type { PptxPresentation } from '../../../src/pptx/types.js';

describe('PPTX serializer - theme', () => {
  it('serializes theme', () => {
    const pres: PptxPresentation = {
      meta: {},
      slides: [],
      masters: [],
      layouts: [],
      theme: {
        colorScheme: {
          name: 'Office',
          colors: {
            dk1: '000000',
            lt1: 'FFFFFF',
            accent1: '4472C4',
          },
        },
        fontScheme: {
          name: 'Office',
          majorFont: 'Calibri',
          minorFont: 'Calibri',
        },
      },
    };

    const result = semanticToXml(pres);
    expect(result.theme).toBeDefined();
    expect(result.theme!).toContain('a:theme');
    expect(result.theme!).toContain('a:clrScheme');
    expect(result.theme!).toContain('name="Office"');
    expect(result.theme!).toContain('a:srgbClr');
    expect(result.theme!).toContain('val="000000"');
    expect(result.theme!).toContain('val="4472C4"');
    expect(result.theme!).toContain('a:fontScheme');
    expect(result.theme!).toContain('a:majorFont');
    expect(result.theme!).toContain('typeface="Calibri"');
  });
});
