import { describe, it, expect } from 'vitest';
import { parseDocxXml } from '../../src/docx/parser.js';

describe('parseDocxXml', () => {
  it('parses a minimal document.xml', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = parseDocxXml(xml);
    expect(result.tag).toBe('w:document');
    expect(result.children).toHaveLength(1);

    const body = result.children[0] as any;
    expect(body.tag).toBe('w:body');

    const para = body.children[0] as any;
    expect(para.tag).toBe('w:p');

    const run = para.children[0] as any;
    expect(run.tag).toBe('w:r');

    const text = run.children[0] as any;
    expect(text.tag).toBe('w:t');
    expect(text.children[0]).toBe('Hello World');
  });

  it('preserves namespace attributes', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body/>
</w:document>`;

    const result = parseDocxXml(xml);
    expect(result.attrs['xmlns:w']).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    expect(result.attrs['xmlns:r']).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  });

  it('handles paragraph with multiple runs', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First</w:t></w:r>
      <w:r><w:t>Second</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = parseDocxXml(xml);
    const body = result.children[0] as any;
    const para = body.children[0] as any;
    expect(para.children).toHaveLength(2);
  });
});
