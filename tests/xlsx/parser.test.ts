import { describe, it, expect } from 'vitest';
import { parseXlsxXml } from '../../src/xlsx/parser.js';

describe('parseXlsxXml', () => {
  it('parses a minimal workbook.xml', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

    const result = parseXlsxXml(xml);
    expect(result.tag).toBe('workbook');
    const sheets = result.children.find((c: any) => c.tag === 'sheets');
    expect(sheets).toBeDefined();
    expect(sheets!.children).toHaveLength(1);
    expect(sheets!.children[0].attrs['name']).toBe('Sheet1');
  });

  it('parses a minimal worksheet.xml', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>`;

    const result = parseXlsxXml(xml);
    expect(result.tag).toBe('worksheet');
    const sheetData = result.children.find((c: any) => c.tag === 'sheetData');
    expect(sheetData).toBeDefined();
    expect(sheetData!.children).toHaveLength(1);
    expect(sheetData!.children[0].tag).toBe('row');
  });

  it('parses shared strings XML', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
</sst>`;

    const result = parseXlsxXml(xml);
    expect(result.tag).toBe('sst');
    expect(result.children).toHaveLength(2);
  });

  it('parses styles XML', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf/>
  </cellXfs>
</styleSheet>`;

    const result = parseXlsxXml(xml);
    expect(result.tag).toBe('styleSheet');
    expect(result.children.find((c: any) => c.tag === 'fonts')).toBeDefined();
    expect(result.children.find((c: any) => c.tag === 'fills')).toBeDefined();
  });
});
