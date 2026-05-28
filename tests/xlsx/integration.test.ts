import { describe, it, expect } from 'vitest';
import {
  parseXlsx, serializeXlsx,
  createXlsx, validateXlsx, loadXlsx, saveXlsx,
  addComment, removeComment, listComments, listSheetComments, getCommentText, updateComment,
  updateXlsxTitle, updateXlsxCreator, updateXlsxSubject, updateXlsxDescription,
  updateXlsxKeywords, updateXlsxCategory, updateXlsxLastModifiedBy,
  xlsx, toXlsxJSON, toXlsxJSONString,
} from '../../src/xlsx/index.js';
import { OMP } from '../../src/omp.js';
import JSZip from 'jszip';
import { join } from 'path';
import { tmpdir } from 'os';

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

describe('XLSX create', () => {
  it('createXlsx returns valid workbook', () => {
    const wb = createXlsx({ title: 'Test', creator: 'Alice', sheetName: 'Data' });
    expect(wb.meta.title).toBe('Test');
    expect(wb.meta.creator).toBe('Alice');
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe('Data');
    expect(wb.sharedStrings).toHaveLength(0);
  });

  it('createXlsx serializes and round-trips', async () => {
    const wb = createXlsx({ title: 'RT' });
    wb.sheets[0].cells.push([{ value: 'A1', type: 'string' }]);
    const buf = await serializeXlsx(wb);
    expect(buf.byteLength).toBeGreaterThan(0);
  });
});

describe('XLSX validate', () => {
  it('validateXlsx passes for valid workbook', () => {
    const wb = createXlsx({ title: 'Valid' });
    const issues = validateXlsx(wb);
    expect(issues.filter(i => i.level === 'error')).toHaveLength(0);
  });
});

describe('XLSX comments', () => {
  it('addComment and listComments', () => {
    const wb = createXlsx({ title: 'Comments' });
    const c = addComment(wb, 0, 'A1', 'Reviewer', 'Check this');
    expect(c.ref).toBe('A1');
    expect(c.text).toBe('Check this');
    expect(wb.sheets[0].comments).toHaveLength(1);

    const list = listComments(wb);
    expect(list).toHaveLength(1);
    expect(list[0].sheetIndex).toBe(0);
  });

  it('listSheetComments filters by sheet', () => {
    const wb = createXlsx({ title: 'Test' });
    addComment(wb, 0, 'A1', 'A', 'comment 1');
    addComment(wb, 0, 'B2', 'B', 'comment 2');
    expect(listSheetComments(wb, 0)).toHaveLength(2);
  });

  it('getCommentText and updateComment', () => {
    const wb = createXlsx({ title: 'Test' });
    addComment(wb, 0, 'A1', 'Author', 'Original');
    expect(getCommentText(wb, 0, 'A1')).toBe('Original');
    expect(updateComment(wb, 0, 'A1', 'Updated')).toBe(true);
    expect(getCommentText(wb, 0, 'A1')).toBe('Updated');
  });

  it('removeComment', () => {
    const wb = createXlsx({ title: 'Test' });
    addComment(wb, 0, 'A1', 'Author', 'To remove');
    expect(removeComment(wb, 0, 'A1')).toBe(true);
    expect(wb.sheets[0].comments).toHaveLength(0);
  });

  it('comments survive round-trip', async () => {
    const wb = createXlsx({ title: 'RT' });
    addComment(wb, 0, 'A1', 'Author', 'Round trip');
    const path = join(tmpdir(), `omp-test-xlsx-comments-${Date.now()}.xlsx`);
    await saveXlsx(wb, path);
    const loaded = await loadXlsx(path);
    expect(loaded.semantic.sheets[0].comments!.length).toBeGreaterThanOrEqual(1);
  });
});

describe('XLSX meta updates', () => {
  it('updateXlsxTitle', () => {
    const wb = createXlsx({ title: 'old' });
    updateXlsxTitle(wb, 'new');
    expect(wb.meta.title).toBe('new');
    expect(wb.meta.modified).toBeDefined();
  });

  it('all update functions work', () => {
    const wb = createXlsx({});
    updateXlsxCreator(wb, 'C');
    updateXlsxSubject(wb, 'S');
    updateXlsxDescription(wb, 'D');
    updateXlsxKeywords(wb, 'K');
    updateXlsxCategory(wb, 'Cat');
    updateXlsxLastModifiedBy(wb, 'LMB');
    expect(wb.meta.creator).toBe('C');
    expect(wb.meta.subject).toBe('S');
    expect(wb.meta.description).toBe('D');
    expect(wb.meta.keywords).toBe('K');
    expect(wb.meta.category).toBe('Cat');
    expect(wb.meta.lastModifiedBy).toBe('LMB');
  });

  it('namespace xlsx.updateTitle', () => {
    const wb = createXlsx({});
    xlsx.updateTitle(wb, 'ns');
    expect(wb.meta.title).toBe('ns');
  });

  it('xlsx.toJSON', () => {
    const wb = createXlsx({ title: 'JSON' });
    const json = toXlsxJSON(wb);
    expect(json.meta.title).toBe('JSON');
    const str = toXlsxJSONString(wb);
    expect(str).toContain('JSON');
    const nsJson = xlsx.toJSON(wb);
    expect(nsJson.meta.title).toBe('JSON');
  });
});

describe('XLSX OMP namespace', () => {
  it('OMP.xlsx.create / updateTitle / validate / serialize', async () => {
    const wb = OMP.xlsx.create({ title: 'OMP' });
    expect(wb.meta.title).toBe('OMP');

    OMP.xlsx.updateTitle(wb, 'Updated');
    expect(wb.meta.title).toBe('Updated');

    const issues = OMP.xlsx.validate(wb);
    expect(Array.isArray(issues)).toBe(true);

    const buf = await OMP.xlsx.serialize(wb);
    expect(buf.byteLength).toBeGreaterThan(0);
  });

  it('OMP.xlsx.addComment / getCommentText', () => {
    const wb = OMP.xlsx.create({ title: 'Test' });
    OMP.xlsx.addComment(wb, 0, 'A1', 'Author', 'Comment');
    expect(OMP.xlsx.getCommentText(wb, 0, 'A1')).toBe('Comment');
    expect(OMP.xlsx.listComments(wb)).toHaveLength(1);
  });
});
