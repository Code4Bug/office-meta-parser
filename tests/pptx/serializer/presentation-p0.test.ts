import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/pptx/serializer.js';
import type { PptxPresentation } from '../../../src/pptx/types.js';

describe('PPTX serializer - presentation P0 features', () => {
  it('serializes custom slideSize', () => {
    const pres: PptxPresentation = {
      meta: {},
      slides: [],
      masters: [],
      layouts: [],
      slideSize: { width: 12192000, height: 6858000 },
    };

    const result = semanticToXml(pres);
    expect(result.presentation).toContain('cx="12192000"');
    expect(result.presentation).toContain('cy="6858000"');
  });

  it('serializes custom notesSize', () => {
    const pres: PptxPresentation = {
      meta: {},
      slides: [],
      masters: [],
      layouts: [],
      notesSize: { width: 6858000, height: 9144000 },
    };

    const result = semanticToXml(pres);
    expect(result.presentation).toContain('cx="6858000"');
    expect(result.presentation).toContain('cy="9144000"');
  });
});
