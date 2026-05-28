import { describe, it, expect } from 'vitest';
import { serializeStyles } from '../../../src/docx/serializer.js';
import type { StyleDefinitions } from '../../../src/docx/types.js';

describe('DOCX serializer - styles', () => {
  it('serializes paragraph styles', () => {
    const styles: StyleDefinitions = {
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'heading 1',
          properties: { alignment: 'left', spacing: { before: 240, after: 120 } },
        },
      ],
      characterStyles: [],
      tableStyles: [],
    };

    const xml = serializeStyles(styles);
    expect(xml).toContain('w:styles');
    expect(xml).toContain('w:style');
    expect(xml).toContain('w:type="paragraph"');
    expect(xml).toContain('w:styleId="Heading1"');
    expect(xml).toContain('w:name');
    expect(xml).toContain('heading 1');
    expect(xml).toContain('w:jc');
    expect(xml).toContain('w:spacing');
  });

  it('serializes character styles', () => {
    const styles: StyleDefinitions = {
      paragraphStyles: [],
      characterStyles: [
        {
          id: 'Strong',
          name: 'strong',
          properties: { bold: true },
        },
      ],
      tableStyles: [],
    };

    const xml = serializeStyles(styles);
    expect(xml).toContain('w:type="character"');
    expect(xml).toContain('w:styleId="Strong"');
    expect(xml).toContain('w:b');
  });

  it('serializes table styles', () => {
    const styles: StyleDefinitions = {
      paragraphStyles: [],
      characterStyles: [],
      tableStyles: [
        {
          id: 'TableGrid',
          name: 'table grid',
        },
      ],
    };

    const xml = serializeStyles(styles);
    expect(xml).toContain('w:type="table"');
    expect(xml).toContain('w:styleId="TableGrid"');
  });
});
