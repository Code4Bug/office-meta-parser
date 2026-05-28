import { describe, it, expect } from 'vitest';
import { serializeHeader, serializeFooter } from '../../../src/docx/serializer.js';
import type { Header, Footer } from '../../../src/docx/types.js';

describe('DOCX serializer - header-footer', () => {
  it('serializes header', () => {
    const header: Header = {
      id: '1',
      type: 'default',
      content: [{ type: 'paragraph', runs: [{ text: 'Header text' }] }],
    };

    const xml = serializeHeader(header);
    expect(xml).toContain('w:hdr');
    expect(xml).toContain('w:p');
    expect(xml).toContain('Header text');
  });

  it('serializes footer', () => {
    const footer: Footer = {
      id: '1',
      type: 'default',
      content: [{ type: 'paragraph', runs: [{ text: 'Footer text' }] }],
    };

    const xml = serializeFooter(footer);
    expect(xml).toContain('w:ftr');
    expect(xml).toContain('w:p');
    expect(xml).toContain('Footer text');
  });
});
