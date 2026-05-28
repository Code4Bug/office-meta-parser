import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../../src/pptx/serializer.js';
import type { PptxPresentation } from '../../../src/pptx/types.js';

describe('PPTX serializer - masters', () => {
  it('serializes slide master', () => {
    const pres: PptxPresentation = {
      meta: {},
      slides: [],
      masters: [
        {
          id: '1',
          layouts: [
            {
              id: '1',
              name: 'Title Slide',
              placeholders: [
                { type: 'ctrTitle', position: { x: 0, y: 0, width: 9144000, height: 1143000 } },
              ],
            },
          ],
          background: { type: 'solid', color: 'FFFFFF' },
        },
      ],
      layouts: [],
    };

    const result = semanticToXml(pres);
    expect(result.masters).toHaveLength(1);
    expect(result.masters[0]).toContain('p:sldMaster');
    expect(result.masters[0]).toContain('p:bg');
    expect(result.masters[0]).toContain('a:solidFill');
    expect(result.masters[0]).toContain('val="FFFFFF"');
    expect(result.masters[0]).toContain('p:sldLayoutId');
  });

  it('serializes slide layout', () => {
    const pres: PptxPresentation = {
      meta: {},
      slides: [],
      masters: [],
      layouts: [
        {
          id: '1',
          name: 'Title Slide',
          placeholders: [
            { type: 'ctrTitle', position: { x: 0, y: 0, width: 9144000, height: 1143000 }, index: 0 },
            { type: 'subTitle', position: { x: 0, y: 1143000, width: 9144000, height: 1143000 }, index: 1 },
          ],
        },
      ],
    };

    const result = semanticToXml(pres);
    expect(result.layouts).toHaveLength(1);
    expect(result.layouts[0]).toContain('p:sldLayout');
    expect(result.layouts[0]).toContain('name="Title Slide"');
    expect(result.layouts[0]).toContain('p:ph');
    expect(result.layouts[0]).toContain('type="ctrTitle"');
    expect(result.layouts[0]).toContain('type="subTitle"');
  });
});
