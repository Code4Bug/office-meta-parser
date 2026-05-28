import { describe, it, expect } from 'vitest';
import { parsePptx, serializePptx } from '../../src/pptx/index.js';
import JSZip from 'jszip';

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
