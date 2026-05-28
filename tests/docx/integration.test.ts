import { describe, it, expect } from 'vitest';
import { parseDocx, serializeDocx, parseDocxXml, serializeDocxXml } from '../../src/docx/index.js';
import JSZip from 'jszip';

async function createMinimalDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats.wordprocessingml.document.main+xml"/>
</Types>`);

  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`);

  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('DOCX integration', () => {
  it('parseDocx returns raw and semantic views', async () => {
    const buffer = await createMinimalDocx();
    const result = await parseDocx(buffer);

    expect(result.raw).toBeDefined();
    expect(result.semantic).toBeDefined();
    expect(result.raw.parts.size).toBeGreaterThan(0);
  });

  it('semantic view has correct structure', async () => {
    const buffer = await createMinimalDocx();
    const { semantic } = await parseDocx(buffer);

    expect(semantic.body.blocks).toHaveLength(1);
    const para = semantic.body.blocks[0] as any;
    expect(para.type).toBe('paragraph');
    expect(para.runs[0].text).toBe('Hello World');
  });

  it('serializeDocx produces valid ZIP', async () => {
    const buffer = await createMinimalDocx();
    const { semantic } = await parseDocx(buffer);

    const output = await serializeDocx(semantic);
    expect(output).toBeInstanceOf(ArrayBuffer);

    const zip = await JSZip.loadAsync(output);
    expect(zip.file('word/document.xml')).not.toBeNull();
  });

  it('parseDocxXml parses XML string', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Test</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = parseDocxXml(xml);
    expect(result.tag).toBe('w:document');
  });

  it('serializeDocxXml produces XML string', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Test</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = serializeDocxXml(parseDocxXml(xml));
    expect(result).toContain('w:document');
    expect(result).toContain('Test');
  });
});
