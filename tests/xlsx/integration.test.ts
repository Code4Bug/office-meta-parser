import { describe, it, expect } from 'vitest';
import { parseXlsx, serializeXlsx } from '../../src/xlsx/index.js';
import JSZip from 'jszip';

async function createMinimalXlsx(): Promise<ArrayBuffer> {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats.spreadsheetml.sharedStrings+xml"/>
</Types>`);

  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);

  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`);

  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);

  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>`);

  zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>`);

  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('XLSX integration', () => {
  it('parseXlsx returns raw and semantic views', async () => {
    const buffer = await createMinimalXlsx();
    const result = await parseXlsx(buffer);

    expect(result.raw).toBeDefined();
    expect(result.semantic).toBeDefined();
    expect(result.raw.parts.size).toBeGreaterThan(0);
  });

  it('semantic view has correct structure', async () => {
    const buffer = await createMinimalXlsx();
    const { semantic } = await parseXlsx(buffer);

    expect(semantic.sheets).toHaveLength(1);
    expect(semantic.sheets[0].name).toBe('Sheet1');
    expect(semantic.sheets[0].cells).toHaveLength(1);
    expect(semantic.sheets[0].cells[0][0].value).toBe('Hello');
    expect(semantic.sheets[0].cells[0][1].value).toBe(42);
  });

  it('serializeXlsx produces valid ZIP', async () => {
    const buffer = await createMinimalXlsx();
    const { semantic } = await parseXlsx(buffer);

    const output = await serializeXlsx(semantic);
    expect(output).toBeInstanceOf(ArrayBuffer);

    const zip = await JSZip.loadAsync(output);
    expect(zip.file('xl/workbook.xml')).not.toBeNull();
    expect(zip.file('xl/worksheets/sheet1.xml')).not.toBeNull();
  });
});
