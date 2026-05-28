import { describe, it, expect } from 'vitest';
import { parseParagraph } from '../../../src/docx/parsers/paragraph.js';
import type { ParsedNode } from '../../../src/core/types.js';

function makeRun(rPrChildren: ParsedNode[]): ParsedNode {
  return {
    tag: 'w:p',
    attrs: {},
    children: [
      {
        tag: 'w:r',
        attrs: {},
        children: [
          { tag: 'w:rPr', attrs: {}, children: rPrChildren },
          { tag: 'w:t', attrs: {}, children: ['test'] },
        ],
      },
    ],
  };
}

describe('DOCX parser - run properties P0', () => {
  it('parses highlight', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:highlight', attrs: { 'w:val': 'yellow' }, children: [] },
    ]));
    expect((para.runs[0] as any).highlight).toBe('yellow');
  });

  it('parses run shading', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:shd', attrs: { 'w:fill': 'FFFF00', 'w:val': 'clear' }, children: [] },
    ]));
    expect((para.runs[0] as any).shadingColor).toBe('FFFF00');
    expect((para.runs[0] as any).shadingPattern).toBe('clear');
  });

  it('parses caps', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:caps', attrs: {}, children: [] },
    ]));
    expect((para.runs[0] as any).caps).toBe(true);
  });

  it('parses smallCaps', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:smallCaps', attrs: {}, children: [] },
    ]));
    expect((para.runs[0] as any).smallCaps).toBe(true);
  });

  it('parses dstrike', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:dstrike', attrs: {}, children: [] },
    ]));
    expect((para.runs[0] as any).dstrike).toBe(true);
  });

  it('parses vanish', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:vanish', attrs: {}, children: [] },
    ]));
    expect((para.runs[0] as any).vanish).toBe(true);
  });

  it('parses characterSpacing', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:spacing', attrs: { 'w:val': '40' }, children: [] },
    ]));
    expect((para.runs[0] as any).characterSpacing).toBe(40);
  });

  it('parses kern', () => {
    const para = parseParagraph(makeRun([
      { tag: 'w:kern', attrs: { 'w:val': '1440' }, children: [] },
    ]));
    expect((para.runs[0] as any).kern).toBe(1440);
  });
});
