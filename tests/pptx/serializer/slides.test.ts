import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/pptx/serializer.js';
import type { PptxPresentation } from '../../../src/pptx/types.js';

function makePresentation(slides: any[]): PptxPresentation {
  return {
    meta: {},
    slides,
    masters: [],
    layouts: [],
  };
}

describe('PPTX serializer - slides', () => {
  it('serializes slide with transition', () => {
    const pres = makePresentation([
      {
        elements: [],
        transition: { type: 'fade', duration: 500 },
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:transition');
    expect(result.slides[0]).toContain('p:fade');
    expect(result.slides[0]).toContain('spd="fast"');
  });

  it('serializes slide with animations', () => {
    const pres = makePresentation([
      {
        elements: [],
        animations: [
          {
            trigger: 'onClick',
            type: 'parallel',
            children: [
              {
                trigger: 'onClick',
                type: 'set',
                shapeId: '1',
                duration: 1000,
              },
            ],
          },
        ],
      },
    ]);

    const result = semanticToXml(pres);
    expect(result.slides[0]).toContain('p:timing');
    expect(result.slides[0]).toContain('p:tnLst');
    expect(result.slides[0]).toContain('p:par');
    expect(result.slides[0]).toContain('p:cTn');
    expect(result.slides[0]).toContain('triggerType="onClick"');
    expect(result.slides[0]).toContain('p:set');
    expect(result.slides[0]).toContain('p:spTgt');
    expect(result.slides[0]).toContain('spid="1"');
  });
});
