import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/pptx/serializer.js';
import type { PptxPresentation } from '../../../src/pptx/types.js';

describe('PPTX serializer - presentation', () => {
  it('serializes a presentation with slides', () => {
    const pres: PptxPresentation = {
      meta: {},
      slides: [
        { elements: [] },
      ],
      masters: [],
      layouts: [],
    };

    const result = semanticToXml(pres);
    expect(result.presentation).toContain('p:presentation');
    expect(result.slides).toHaveLength(1);
  });

  it('generates relationship files', () => {
    const pres: PptxPresentation = {
      meta: {},
      slides: [
        { elements: [] },
      ],
      masters: [],
      layouts: [],
    };

    const result = semanticToXml(pres);
    expect(result.rels).toContain('Relationships');
    expect(result.slideRels).toHaveLength(1);
  });
});
