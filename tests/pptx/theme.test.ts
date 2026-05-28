import { describe, it, expect } from 'vitest';
import { extractTheme } from '../../src/pptx/parsers/theme.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('extractTheme', () => {
  it('returns undefined when no theme file', () => {
    const raw = makeRaw({});
    expect(extractTheme(raw)).toBeUndefined();
  });

  it('parses color scheme', () => {
    const themeXml: ParsedNode = {
      tag: 'a:theme',
      attrs: {},
      children: [
        {
          tag: 'a:themeElements',
          attrs: {},
          children: [
            {
              tag: 'a:clrScheme',
              attrs: { name: 'Office' },
              children: [
                {
                  tag: 'a:dk1',
                  attrs: {},
                  children: [{ tag: 'a:srgbClr', attrs: { val: '000000' }, children: [] }],
                },
                {
                  tag: 'a:lt1',
                  attrs: {},
                  children: [{ tag: 'a:srgbClr', attrs: { val: 'FFFFFF' }, children: [] }],
                },
                {
                  tag: 'a:accent1',
                  attrs: {},
                  children: [{ tag: 'a:srgbClr', attrs: { val: '4472C4' }, children: [] }],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'ppt/theme/theme1.xml': themeXml });
    const theme = extractTheme(raw);
    expect(theme).toBeDefined();
    expect(theme!.colorScheme.name).toBe('Office');
    expect(theme!.colorScheme.colors['dk1']).toBe('000000');
    expect(theme!.colorScheme.colors['lt1']).toBe('FFFFFF');
    expect(theme!.colorScheme.colors['accent1']).toBe('4472C4');
  });

  it('parses font scheme', () => {
    const themeXml: ParsedNode = {
      tag: 'a:theme',
      attrs: {},
      children: [
        {
          tag: 'a:themeElements',
          attrs: {},
          children: [
            {
              tag: 'a:fontScheme',
              attrs: { name: 'Office' },
              children: [
                {
                  tag: 'a:majorFont',
                  attrs: {},
                  children: [
                    { tag: 'a:latin', attrs: { typeface: 'Calibri Light' }, children: [] },
                  ],
                },
                {
                  tag: 'a:minorFont',
                  attrs: {},
                  children: [
                    { tag: 'a:latin', attrs: { typeface: 'Calibri' }, children: [] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'ppt/theme/theme1.xml': themeXml });
    const theme = extractTheme(raw);
    expect(theme).toBeDefined();
    expect(theme!.fontScheme.name).toBe('Office');
    expect(theme!.fontScheme.majorFont).toBe('Calibri Light');
    expect(theme!.fontScheme.minorFont).toBe('Calibri');
  });

  it('parses complete theme', () => {
    const themeXml: ParsedNode = {
      tag: 'a:theme',
      attrs: {},
      children: [
        {
          tag: 'a:themeElements',
          attrs: {},
          children: [
            {
              tag: 'a:clrScheme',
              attrs: { name: 'Custom' },
              children: [
                {
                  tag: 'a:dk1',
                  attrs: {},
                  children: [{ tag: 'a:srgbClr', attrs: { val: '333333' }, children: [] }],
                },
                {
                  tag: 'a:hlink',
                  attrs: {},
                  children: [{ tag: 'a:srgbClr', attrs: { val: '0563C1' }, children: [] }],
                },
              ],
            },
            {
              tag: 'a:fontScheme',
              attrs: { name: 'Custom' },
              children: [
                {
                  tag: 'a:majorFont',
                  attrs: {},
                  children: [
                    { tag: 'a:latin', attrs: { typeface: 'Arial' }, children: [] },
                  ],
                },
                {
                  tag: 'a:minorFont',
                  attrs: {},
                  children: [
                    { tag: 'a:latin', attrs: { typeface: 'Verdana' }, children: [] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'ppt/theme/theme1.xml': themeXml });
    const theme = extractTheme(raw);
    expect(theme).toBeDefined();
    expect(theme!.colorScheme.colors['dk1']).toBe('333333');
    expect(theme!.colorScheme.colors['hlink']).toBe('0563C1');
    expect(theme!.fontScheme.majorFont).toBe('Arial');
    expect(theme!.fontScheme.minorFont).toBe('Verdana');
  });
});
