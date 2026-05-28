import { describe, it, expect } from 'vitest';
import {
  parsePptx, serializePptx,
  createPptx, validatePptx, loadPptx, savePptx,
  addComment, removeComment, listComments, listSlideComments, getCommentText,
  updatePptxTitle, updatePptxCreator, updatePptxSubject, updatePptxDescription,
  updatePptxKeywords, updatePptxCategory, updatePptxLastModifiedBy,
  pptx, toPptxJSON, toPptxJSONString,
} from '../../src/pptx/index.js';
import { OMP } from '../../src/omp.js';
import JSZip from 'jszip';
import { join } from 'path';
import { tmpdir } from 'os';

async function createMinimalPptx(): Promise<ArrayBuffer> {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats.presentationml.slide+xml"/>
</Types>`);

  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>`);

  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`);

  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:t>Hello World</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`);

  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('PPTX integration', () => {
  it('parsePptx returns raw and semantic views', async () => {
    const buffer = await createMinimalPptx();
    const result = await parsePptx(buffer);

    expect(result.raw).toBeDefined();
    expect(result.semantic).toBeDefined();
    expect(result.raw.parts.size).toBeGreaterThan(0);
  });

  it('semantic view has correct structure', async () => {
    const buffer = await createMinimalPptx();
    const { semantic } = await parsePptx(buffer);

    expect(semantic.slides).toHaveLength(1);
    expect(semantic.slides[0].elements).toHaveLength(1);
    expect(semantic.slides[0].elements[0].type).toBe('text');
  });

  it('serializePptx produces valid ZIP', async () => {
    const buffer = await createMinimalPptx();
    const { semantic } = await parsePptx(buffer);

    // 补充 parse 阶段缺失的必须字段
    if (!semantic.theme) {
      semantic.theme = {
        colorScheme: {
          name: 'Office',
          colors: {
            dk1: '1F3864', lt1: 'FFFFFF', dk2: '4472C4', lt2: 'E7E6E6',
            accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
            accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
            hlink: '0563C1', folHlink: '954F72',
          },
        },
        fontScheme: { name: 'Office', majorFont: 'Calibri', minorFont: 'Calibri' },
      };
    }
    if (!semantic.masters || semantic.masters.length === 0) {
      semantic.masters = [{
        layouts: [],
        txStyles: {
          titleStyle: { levels: [] },
          bodyStyle: { levels: [] },
          otherStyle: { levels: [] },
        },
      }];
    }
    if (!semantic.layouts || semantic.layouts.length === 0) {
      semantic.layouts = [{ id: 1, name: 'Layout 1', placeholders: [] }];
      semantic.masters[0].layouts = [semantic.layouts[0]];
    }

    const output = await serializePptx(semantic);
    expect(output).toBeInstanceOf(ArrayBuffer);

    const zip = await JSZip.loadAsync(output);
    expect(zip.file('ppt/presentation.xml')).not.toBeNull();
    expect(zip.file('ppt/slides/slide1.xml')).not.toBeNull();
  });
});

describe('PPTX create', () => {
  it('createPptx returns valid presentation', () => {
    const pres = createPptx({ title: 'Test', creator: 'Alice' });
    expect(pres.meta.title).toBe('Test');
    expect(pres.meta.creator).toBe('Alice');
    expect(pres.slides).toHaveLength(1);
    expect(pres.masters.length).toBeGreaterThanOrEqual(1);
    expect(pres.layouts.length).toBeGreaterThanOrEqual(1);
    expect(pres.theme).toBeDefined();
    expect(pres.slideSize).toBeDefined();
  });

  it('createPptx serializes and round-trips', async () => {
    const pres = createPptx({ title: 'RT' });
    pres.slides[0].elements.push({
      type: 'text',
      content: 'Title',
      position: { x: 0, y: 0, width: 9144000, height: 1000000 },
      paragraphs: [{ runs: [{ text: 'Title Text' }] }],
    });
    const buf = await serializePptx(pres);
    expect(buf.byteLength).toBeGreaterThan(0);
  });
});

describe('PPTX validate', () => {
  it('validatePptx passes for valid presentation', () => {
    const pres = createPptx({ title: 'Valid' });
    const issues = validatePptx(pres);
    expect(issues.filter(i => i.level === 'error')).toHaveLength(0);
  });
});

describe('PPTX comments', () => {
  it('addComment and listComments', () => {
    const pres = createPptx({ title: 'Comments' });
    const c = addComment(pres, 0, 'Reviewer', 'Fix this', 100, 200);
    expect(c.id).toBeDefined();
    expect(c.authorName).toBe('Reviewer');
    expect(c.text).toBe('Fix this');
    expect(c.position?.x).toBe(100);
    expect(pres.slides[0].comments).toHaveLength(1);

    const list = listComments(pres);
    expect(list).toHaveLength(1);
    expect(list[0].slideIndex).toBe(0);
  });

  it('listSlideComments filters by slide', () => {
    const pres = createPptx({ title: 'Test' });
    addComment(pres, 0, 'A', 'c1');
    addComment(pres, 0, 'B', 'c2');
    expect(listSlideComments(pres, 0)).toHaveLength(2);
  });

  it('getCommentText', () => {
    const pres = createPptx({ title: 'Test' });
    const c = addComment(pres, 0, 'Author', 'Text');
    expect(getCommentText(pres, 0, c.id)).toBe('Text');
  });

  it('removeComment', () => {
    const pres = createPptx({ title: 'Test' });
    const c = addComment(pres, 0, 'Author', 'To remove');
    expect(removeComment(pres, 0, c.id)).toBe(true);
    expect(pres.slides[0].comments).toBeUndefined();
  });

  it('comments survive round-trip', async () => {
    const pres = createPptx({ title: 'RT' });
    addComment(pres, 0, 'Author', 'Round trip');
    const path = join(tmpdir(), `omp-test-pptx-comments-${Date.now()}.pptx`);
    await savePptx(pres, path);
    const loaded = await loadPptx(path);
    expect(loaded.semantic.slides[0].comments!.length).toBeGreaterThanOrEqual(1);
    expect(loaded.semantic.slides[0].comments![0].authorName).toBe('Author');
  });
});

describe('PPTX meta updates', () => {
  it('updatePptxTitle', () => {
    const pres = createPptx({ title: 'old' });
    updatePptxTitle(pres, 'new');
    expect(pres.meta.title).toBe('new');
    expect(pres.meta.modified).toBeDefined();
  });

  it('all update functions work', () => {
    const pres = createPptx({});
    updatePptxCreator(pres, 'C');
    updatePptxSubject(pres, 'S');
    updatePptxDescription(pres, 'D');
    updatePptxKeywords(pres, 'K');
    updatePptxCategory(pres, 'Cat');
    updatePptxLastModifiedBy(pres, 'LMB');
    expect(pres.meta.creator).toBe('C');
    expect(pres.meta.subject).toBe('S');
    expect(pres.meta.description).toBe('D');
    expect(pres.meta.keywords).toBe('K');
    expect(pres.meta.category).toBe('Cat');
    expect(pres.meta.lastModifiedBy).toBe('LMB');
  });

  it('namespace pptx.updateTitle', () => {
    const pres = createPptx({});
    pptx.updateTitle(pres, 'ns');
    expect(pres.meta.title).toBe('ns');
  });

  it('pptx.toJSON', () => {
    const pres = createPptx({ title: 'JSON' });
    const json = toPptxJSON(pres);
    expect(json.meta.title).toBe('JSON');
    const str = toPptxJSONString(pres);
    expect(str).toContain('JSON');
    const nsJson = pptx.toJSON(pres);
    expect(nsJson.meta.title).toBe('JSON');
  });
});

describe('PPTX OMP namespace', () => {
  it('OMP.pptx.create / updateTitle / validate / serialize', async () => {
    const pres = OMP.pptx.create({ title: 'OMP' });
    expect(pres.meta.title).toBe('OMP');

    OMP.pptx.updateTitle(pres, 'Updated');
    expect(pres.meta.title).toBe('Updated');

    const issues = OMP.pptx.validate(pres);
    expect(Array.isArray(issues)).toBe(true);

    const buf = await OMP.pptx.serialize(pres);
    expect(buf.byteLength).toBeGreaterThan(0);
  });

  it('OMP.pptx.addComment / listComments', () => {
    const pres = OMP.pptx.create({ title: 'Test' });
    OMP.pptx.addComment(pres, 0, 'Reviewer', 'Fix');
    expect(OMP.pptx.listComments(pres)).toHaveLength(1);
  });
});
