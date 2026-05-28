import { describe, it, expect } from 'vitest';
import { parsePptxXml } from '../../src/pptx/parser.js';

describe('parsePptxXml', () => {
  it('parses a minimal presentation.xml', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>`;

    const result = parsePptxXml(xml);
    expect(result.tag).toBe('p:presentation');
    expect(result.children.find((c: any) => c.tag === 'p:sldMasterIdLst')).toBeDefined();
    expect(result.children.find((c: any) => c.tag === 'p:sldIdLst')).toBeDefined();
  });

  it('parses a minimal slide.xml', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</p:sld>`;

    const result = parsePptxXml(xml);
    expect(result.tag).toBe('p:sld');
    const cSld = result.children.find((c: any) => c.tag === 'p:cSld');
    expect(cSld).toBeDefined();
  });

  it('parses slide layout', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="title">
  <p:cSld name="Title Slide"/>
</p:sldLayout>`;

    const result = parsePptxXml(xml);
    expect(result.tag).toBe('p:sldLayout');
    expect(result.attrs['type']).toBe('title');
  });
});
