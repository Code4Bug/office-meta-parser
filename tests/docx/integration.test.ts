import { describe, it, expect } from 'vitest';
import {
  parseDocx, serializeDocx, parseDocxXml, serializeDocxXml,
  createDocx, validateDocx, loadDocx, saveDocx,
  addComment, removeComment, listComments, getCommentText, markCommentDone, markCommentUndone,
  markInsert, markDelete, addFormatChange, clearRevision,
  listRevisions, hasPendingRevisions, acceptAllInserts, acceptAllDeletes,
  rejectAllInserts, rejectAllDeletes,
  updateDocxTitle, updateDocxCreator, updateDocxSubject, updateDocxDescription,
  updateDocxKeywords, updateDocxCategory, updateDocxLastModifiedBy,
  docx, toDocxJSON, toDocxJSONString,
} from '../../src/docx/index.js';
import { OMP } from '../../src/omp.js';
import JSZip from 'jszip';
import { join } from 'path';
import { tmpdir } from 'os';

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

describe('DOCX create', () => {
  it('createDocx returns valid document', () => {
    const doc = createDocx({ title: 'Test', creator: 'Alice' });
    expect(doc.meta.title).toBe('Test');
    expect(doc.meta.creator).toBe('Alice');
    expect(doc.meta.created).toBeDefined();
    expect(doc.body.blocks).toHaveLength(0);
    expect(doc.styles.paragraphStyles).toHaveLength(0);
  });

  it('createDocx serializes and round-trips', async () => {
    const doc = createDocx({ title: 'Round Trip' });
    doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'Hello' }] });
    const buf = await serializeDocx(doc);
    expect(buf.byteLength).toBeGreaterThan(0);
  });
});

describe('DOCX validate', () => {
  it('validateDocx passes for valid doc', () => {
    const doc = createDocx({ title: 'Valid' });
    doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'Content' }] });
    const issues = validateDocx(doc);
    expect(issues.filter(i => i.level === 'error')).toHaveLength(0);
  });
});

describe('DOCX comments', () => {
  it('addComment and listComments', () => {
    const doc = createDocx({ title: 'Comments Test' });
    const run = { text: 'paragraph text' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });

    const c = addComment(doc, run, 'Reviewer', 'Please fix this');
    expect(c.id).toBeDefined();
    expect(run.commentId).toBe(c.id);
    expect(doc.comments).toHaveLength(1);

    const list = listComments(doc);
    expect(list).toHaveLength(1);
    expect(list[0].comment.author).toBe('Reviewer');
  });

  it('getCommentText returns correct text', () => {
    const doc = createDocx({ title: 'Test' });
    const run = { text: 'text' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });
    const c = addComment(doc, run, 'Author', 'Comment text');
    expect(getCommentText(doc, c.id)).toBe('Comment text');
  });

  it('removeComment clears run reference', () => {
    const doc = createDocx({ title: 'Test' });
    const run = { text: 'text' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });
    const c = addComment(doc, run, 'Author', 'To remove');
    expect(removeComment(doc, c.id)).toBe(true);
    expect(doc.comments).toHaveLength(0);
    expect(run.commentId).toBeUndefined();
  });

  it('markCommentDone / markCommentUndone', () => {
    const doc = createDocx({ title: 'Test' });
    const run = { text: 'text' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });
    const c = addComment(doc, run, 'Author', 'Review');
    expect(markCommentDone(doc, c.id)).toBe(true);
    expect(doc.commentExts?.some(e => e.paraId === c.id && e.done)).toBe(true);
    expect(markCommentUndone(doc, c.id)).toBe(true);
    expect(doc.commentExts?.some(e => e.paraId === c.id && !e.done)).toBe(true);
  });

  it('comments survive round-trip', async () => {
    const doc = createDocx({ title: 'RT' });
    const run = { text: 'content' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });
    addComment(doc, run, 'Author', 'Round trip comment');

    const path = join(tmpdir(), `omp-test-comments-${Date.now()}.docx`);
    await saveDocx(doc, path);
    const loaded = await loadDocx(path);
    expect(loaded.semantic.comments!.length).toBeGreaterThanOrEqual(1);
  });
});

describe('DOCX revisions', () => {
  it('markInsert and markDelete', () => {
    const doc = createDocx({ title: 'Revisions' });
    const run1 = { text: 'inserted' };
    const run2 = { text: 'deleted' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run1, run2] });

    markInsert(run1, 'Alice');
    expect(run1.revisionType).toBe('insert');
    expect(run1.revisionAuthor).toBe('Alice');

    markDelete(run2, 'Bob');
    expect(run2.revisionType).toBe('delete');
  });

  it('hasPendingRevisions and listRevisions', () => {
    const doc = createDocx({ title: 'Test' });
    const run = { text: 'new text' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });
    markInsert(run, 'Author');

    expect(hasPendingRevisions(doc)).toBe(true);
    expect(listRevisions(doc)).toHaveLength(1);
  });

  it('clearRevision removes marks', () => {
    const run = { text: 'text' };
    markInsert(run, 'Author');
    clearRevision(run);
    expect(run.revisionType).toBeUndefined();
    expect(run.revisionAuthor).toBeUndefined();
  });

  it('acceptAllInserts removes insert marks', () => {
    const doc = createDocx({ title: 'Test' });
    const run = { text: 'accepted' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });
    markInsert(run, 'Author');
    const count = acceptAllInserts(doc);
    expect(count).toBeGreaterThanOrEqual(1);
    expect(run.revisionType).toBeUndefined();
  });

  it('acceptAllDeletes removes deleted runs', () => {
    const doc = createDocx({ title: 'Test' });
    const run1 = { text: 'keep' };
    const run2 = { text: 'remove' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run1, run2] });
    markDelete(run2, 'Author');
    acceptAllDeletes(doc);
    expect((doc.body.blocks[0] as any).runs).toHaveLength(1);
  });
});

describe('DOCX meta updates', () => {
  it('updateDocxTitle', () => {
    const doc = createDocx({ title: 'old' });
    updateDocxTitle(doc, 'new');
    expect(doc.meta.title).toBe('new');
    expect(doc.meta.modified).toBeDefined();
  });

  it('all update functions work', () => {
    const doc = createDocx({});
    updateDocxCreator(doc, 'C');
    updateDocxSubject(doc, 'S');
    updateDocxDescription(doc, 'D');
    updateDocxKeywords(doc, 'K');
    updateDocxCategory(doc, 'Cat');
    updateDocxLastModifiedBy(doc, 'LMB');
    expect(doc.meta.creator).toBe('C');
    expect(doc.meta.subject).toBe('S');
    expect(doc.meta.description).toBe('D');
    expect(doc.meta.keywords).toBe('K');
    expect(doc.meta.category).toBe('Cat');
    expect(doc.meta.lastModifiedBy).toBe('LMB');
  });

  it('namespace docx.updateTitle', () => {
    const doc = createDocx({});
    docx.updateTitle(doc, 'ns');
    expect(doc.meta.title).toBe('ns');
  });

  it('docx.toJSON', () => {
    const doc = createDocx({ title: 'JSON' });
    doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'x' }] });
    const json = toDocxJSON(doc);
    expect(json.meta.title).toBe('JSON');
    const str = toDocxJSONString(doc);
    expect(str).toContain('JSON');
    const nsJson = docx.toJSON(doc);
    expect(nsJson.meta.title).toBe('JSON');
  });
});

describe('DOCX OMP namespace', () => {
  it('OMP.docx.create / updateTitle / validate / serialize', async () => {
    const doc = OMP.docx.create({ title: 'OMP' });
    expect(doc.meta.title).toBe('OMP');

    OMP.docx.updateTitle(doc, 'Updated');
    expect(doc.meta.title).toBe('Updated');

    doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'Content' }] });
    const issues = OMP.docx.validate(doc);
    expect(Array.isArray(issues)).toBe(true);

    const buf = await OMP.docx.serialize(doc);
    expect(buf.byteLength).toBeGreaterThan(0);
  });

  it('OMP.docx.addComment', () => {
    const doc = OMP.docx.create({ title: 'Test' });
    const run = { text: 'text' };
    doc.body.blocks.push({ type: 'paragraph', runs: [run] });
    OMP.docx.addComment(doc, run, 'Reviewer', 'Fix');
    expect(OMP.docx.listComments(doc)).toHaveLength(1);
  });

  it('OMP.detectFormat recognizes docx', async () => {
    const doc = OMP.docx.create({ title: 'Test' });
    doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'x' }] });
    const buf = await OMP.docx.serialize(doc);
    const format = await OMP.detectFormat(buf);
    expect(format).toBe('docx');
  });
});
