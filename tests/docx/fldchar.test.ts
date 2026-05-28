import { describe, it, expect } from 'vitest';
import { parseParagraph } from '../../src/docx/parsers/paragraph.js';
import type { ParsedNode } from '../../src/core/types.js';

function fldCharRun(type: string): ParsedNode {
  return {
    tag: 'w:r', attrs: {},
    children: [{ tag: 'w:fldChar', attrs: { 'w:fldCharType': type }, children: [] }],
  };
}

function instrRun(text: string): ParsedNode {
  return {
    tag: 'w:r', attrs: {},
    children: [{ tag: 'w:instrText', attrs: {}, children: [text] }],
  };
}

function textRun(text: string): ParsedNode {
  return {
    tag: 'w:r', attrs: {},
    children: [{ tag: 'w:t', attrs: {}, children: [text] }],
  };
}

describe('DOCX fldChar cross-run', () => {
  it('handles field with instruction split across multiple runs', () => {
    const para: ParsedNode = {
      tag: 'w:p', attrs: {},
      children: [
        fldCharRun('begin'),
        instrRun(' TOC '),
        instrRun('\\o "1-3" '),
        fldCharRun('separate'),
        textRun('Table of Contents'),
        fldCharRun('end'),
      ],
    };
    const result = parseParagraph(para);
    expect(result.runs).toHaveLength(1);
    expect(result.runs[0].field?.instruction).toBe('TOC \\o "1-3"');
    expect(result.runs[0].field?.result).toBe('Table of Contents');
  });

  it('handles field with result split across multiple runs', () => {
    const para: ParsedNode = {
      tag: 'w:p', attrs: {},
      children: [
        fldCharRun('begin'),
        instrRun(' PAGE '),
        fldCharRun('separate'),
        textRun('1'),
        textRun('2'),
        fldCharRun('end'),
      ],
    };
    const result = parseParagraph(para);
    expect(result.runs).toHaveLength(1);
    expect(result.runs[0].field?.instruction).toBe('PAGE');
    expect(result.runs[0].field?.result).toBe('12');
  });

  it('handles field with no separate (no display result)', () => {
    const para: ParsedNode = {
      tag: 'w:p', attrs: {},
      children: [
        fldCharRun('begin'),
        instrRun(' PAGENUM '),
        fldCharRun('end'),
      ],
    };
    const result = parseParagraph(para);
    expect(result.runs).toHaveLength(1);
    expect(result.runs[0].field?.instruction).toBe('PAGENUM');
    expect(result.runs[0].field?.result).toBe('');
  });
});
