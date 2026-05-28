import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../../src/docx/semantic.js';
import type { RawDocument, ParsedNode } from '../../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('DOCX semantic - settings', () => {
  it('extracts settings from settings.xml', () => {
    const settingsXml: ParsedNode = {
      tag: 'w:settings',
      attrs: { 'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' },
      children: [
        { tag: 'w:zoom', attrs: { 'w:percent': '145' }, children: [] },
        { tag: 'w:defaultTabStop', attrs: { 'w:val': '709' }, children: [] },
        { tag: 'w:evenAndOddHeaders', attrs: { 'w:val': '1' }, children: [] },
        { tag: 'w:documentProtection', attrs: { 'w:enforcement': '1' }, children: [] },
        { tag: 'w:characterSpacingControl', attrs: { 'w:val': 'compressPunctuation' }, children: [] },
        {
          tag: 'w:compat',
          attrs: {},
          children: [{
            tag: 'w:compatSetting',
            attrs: {
              'w:name': 'compatibilityMode',
              'w:uri': 'http://schemas.microsoft.com/office/word',
              'w:val': '15',
            },
            children: [],
          }],
        },
      ],
    };

    const raw = makeRaw({ 'word/settings.xml': settingsXml });
    const semantic = rawToSemantic(raw);

    expect(semantic.settings).toBeDefined();
    expect(semantic.settings!.defaultTabStop).toBe(709);
    expect(semantic.settings!.zoom).toBe(145);
    expect(semantic.settings!.compatibilityMode).toBe(15);
    expect(semantic.settings!.evenAndOddHeaders).toBe(true);
    expect(semantic.settings!.documentProtection).toBe(true);
    expect(semantic.settings!.characterSpacingControl).toBe('compressPunctuation');
  });

  it('returns undefined when no settings present', () => {
    const raw = makeRaw({});
    const semantic = rawToSemantic(raw);
    expect(semantic.settings).toBeUndefined();
  });
});
