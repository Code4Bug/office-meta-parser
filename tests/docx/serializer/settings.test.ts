import { describe, it, expect } from 'vitest';
import { serializeSettings } from '../../../src/docx/serializer.js';
import type { DocumentSettings } from '../../../src/docx/types.js';

describe('DOCX serializer - settings', () => {
  it('serializes all settings fields', () => {
    const settings: DocumentSettings = {
      defaultTabStop: 720,
      zoom: 100,
      compatibilityMode: 15,
      evenAndOddHeaders: true,
      documentProtection: true,
      characterSpacingControl: 'compressPunctuation',
    };

    const xml = serializeSettings(settings);
    expect(xml).toContain('w:settings');
    expect(xml).toContain('w:defaultTabStop');
    expect(xml).toContain('w:val="720"');
    expect(xml).toContain('w:percent="100"');
    expect(xml).toContain('w:compat');
    expect(xml).toContain('w:name="compatibilityMode"');
    expect(xml).toContain('w:val="15"');
    expect(xml).toContain('w:evenAndOddHeaders');
    expect(xml).toContain('w:documentProtection');
    expect(xml).toContain('w:val="compressPunctuation"');
  });

  it('serializes partial settings', () => {
    const settings: DocumentSettings = { defaultTabStop: 400 };
    const xml = serializeSettings(settings);
    expect(xml).toContain('w:defaultTabStop');
    expect(xml).toContain('w:val="400"');
    expect(xml).not.toContain('w:zoom');
    expect(xml).not.toContain('w:compat');
  });
});
