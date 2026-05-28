import { unzip, zip } from '../core/zip.js';
import { parseRels } from '../core/rels.js';
import { parseContentTypes } from '../core/content-type.js';
import { parseXml, escapeXmlAttr } from '../core/xml.js';
import { serializeMeta } from '../core/meta.js';
import { serializeAppMeta } from '../core/app-meta.js';
import { serializeCustomProperties } from '../core/custom-meta.js';
import { throwOnError, warnOnWarning } from '../core/validate.js';
import type { RawDocument, ParseResult, ParsedNode } from '../core/types.js';
import type { PptxPresentation, Theme, SlideLayout, SlideMaster } from './types.js';
import { rawToSemantic } from './semantic.js';
import { semanticToXml } from './serializer.js';
import { validatePptx } from './validate.js';

export type {
  PptxPresentation,
  Slide,
  SlideElement,
  SlideComment,
  TextShape,
  ImageShape,
  GroupShape,
  Position,
  ShapeStyle,
  FillStyle,
  BorderStyle,
  ShadowStyle,
  SlideMaster,
  SlideLayout,
  Placeholder,
  Transition,
  Theme,
  ColorScheme,
  FontScheme,
} from './types.js';

export type { AppMeta } from '../core/app-meta.js';
export type { CustomProperty } from '../core/custom-meta.js';

export { parsePptxXml } from './parser.js';
export { rawToSemantic } from './semantic.js';
export { semanticToXml } from './serializer.js';
export { validatePptx } from './validate.js';
export {
  updatePptxTitle, updatePptxSubject, updatePptxCreator,
  updatePptxDescription, updatePptxKeywords, updatePptxLastModifiedBy,
  updatePptxCategory, pptx,
  toPptxJSON, toPptxJSONString, savePptxJSON,
} from './meta.js';
export {
  addComment, removeComment, listComments, listSlideComments, getCommentText,
} from './comments.js';
export type { CommentInfo } from './comments.js';

export async function parsePptx(buffer: ArrayBuffer): Promise<ParseResult<PptxPresentation>> {
  const entries = await unzip(buffer);

  const rels = new Map<string, import('../core/types.js').Relationship[]>();
  const parts = new Map<string, ParsedNode>();
  let contentTypes: import('../core/types.js').ContentType[] = [];

  for (const entry of entries) {
    const text = new TextDecoder().decode(entry.data);

    if (entry.path === '[Content_Types].xml') {
      contentTypes = parseContentTypes(text);
    } else if (entry.path.endsWith('.rels')) {
      rels.set(entry.path, parseRels(text));
    } else if (entry.path.endsWith('.xml')) {
      parts.set(entry.path, parseXml(text));
    }
  }

  const raw: RawDocument = { entries, rels, contentTypes, parts };
  const semantic = rawToSemantic(raw);

  return { raw, semantic };
}

export async function serializePptx(pres: PptxPresentation): Promise<ArrayBuffer> {
  const issues = validatePptx(pres);
  warnOnWarning(issues);
  throwOnError(issues);

  const parts = semanticToXml(pres);
  const raw = pres.rawXmlParts;
  const useRaw = raw && raw.size > 0;

  const metaXml = useRaw && raw!.has('docProps/core.xml') ? raw!.get('docProps/core.xml')! : serializeMeta(pres.meta);

  const entries: { path: string; data: ArrayBuffer }[] = [];

  // Content_Types
  if (useRaw && raw!.has('[Content_Types].xml')) {
    entries.push({ path: '[Content_Types].xml', data: new TextEncoder().encode(raw!.get('[Content_Types].xml')!).buffer });
  } else {
    entries.push({ path: '[Content_Types].xml', data: new TextEncoder().encode(parts.contentTypes).buffer });
  }

  // _rels/.rels
  if (useRaw && raw!.has('_rels/.rels')) {
    entries.push({ path: '_rels/.rels', data: new TextEncoder().encode(raw!.get('_rels/.rels')!).buffer });
  } else {
    entries.push({ path: '_rels/.rels', data: new TextEncoder().encode(buildPackageRels(pres)).buffer });
  }

  entries.push(
    { path: 'docProps/core.xml', data: new TextEncoder().encode(metaXml).buffer },
    { path: 'docProps/app.xml', data: new TextEncoder().encode(pres.appMeta ? serializeAppMeta(pres.appMeta) : buildAppXml(pres)).buffer },
  );

  // ppt/presentation.xml
  if (useRaw && raw!.has('ppt/presentation.xml')) {
    entries.push({ path: 'ppt/presentation.xml', data: new TextEncoder().encode(raw!.get('ppt/presentation.xml')!).buffer });
  } else {
    entries.push({ path: 'ppt/presentation.xml', data: new TextEncoder().encode(parts.presentation).buffer });
  }

  // ppt/_rels/presentation.xml.rels
  if (useRaw && raw!.has('ppt/_rels/presentation.xml.rels')) {
    entries.push({ path: 'ppt/_rels/presentation.xml.rels', data: new TextEncoder().encode(raw!.get('ppt/_rels/presentation.xml.rels')!).buffer });
  } else {
    entries.push({ path: 'ppt/_rels/presentation.xml.rels', data: new TextEncoder().encode(parts.rels).buffer });
  }

  // ppt/presProps.xml
  if (useRaw && raw!.has('ppt/presProps.xml')) {
    entries.push({ path: 'ppt/presProps.xml', data: new TextEncoder().encode(raw!.get('ppt/presProps.xml')!).buffer });
  } else {
    entries.push({ path: 'ppt/presProps.xml', data: new TextEncoder().encode(buildPresProps()).buffer });
  }

  // 幻灯片
  for (let i = 0; i < parts.slides.length; i++) {
    const path = `ppt/slides/slide${i + 1}.xml`;
    if (useRaw && raw!.has(path)) {
      entries.push({ path, data: new TextEncoder().encode(raw!.get(path)!).buffer });
    } else {
      entries.push({ path, data: new TextEncoder().encode(parts.slides[i]).buffer });
    }
  }

  // 幻灯片 rels
  for (let i = 0; i < parts.slideRels.length; i++) {
    const path = `ppt/slides/_rels/slide${i + 1}.xml.rels`;
    if (useRaw && raw!.has(path)) {
      entries.push({ path, data: new TextEncoder().encode(raw!.get(path)!).buffer });
    } else {
      entries.push({ path, data: new TextEncoder().encode(parts.slideRels[i]).buffer });
    }
  }

  // 批注
  for (const c of parts.comments) {
    if (useRaw && raw!.has(c.path)) {
      entries.push({ path: c.path, data: new TextEncoder().encode(raw!.get(c.path)!).buffer });
    } else {
      entries.push({ path: c.path, data: new TextEncoder().encode(c.xml).buffer });
    }
  }

  // 备注幻灯片（notesSlide）从 raw 原样保留
  if (useRaw) {
    for (const [path, xml] of raw!) {
      if (path.startsWith('ppt/notesSlides/') && (path.endsWith('.xml') || path.endsWith('.rels'))) {
        entries.push({ path, data: new TextEncoder().encode(xml).buffer });
      }
    }
  }

  if (pres.customProperties && pres.customProperties.length > 0) {
    const xml = useRaw && raw!.has('docProps/custom.xml') ? raw!.get('docProps/custom.xml')! : serializeCustomProperties(pres.customProperties);
    entries.push({ path: 'docProps/custom.xml', data: new TextEncoder().encode(xml).buffer });
  }

  // 主题
  if (parts.theme) {
    const path = 'ppt/theme/theme1.xml';
    if (useRaw && raw!.has(path)) {
      entries.push({ path, data: new TextEncoder().encode(raw!.get(path)!).buffer });
    } else {
      entries.push({ path, data: new TextEncoder().encode(parts.theme).buffer });
    }
  }

  // 幻灯片母版
  for (let i = 0; i < parts.masters.length; i++) {
    const masterPath = `ppt/slideMasters/slideMaster${i + 1}.xml`;
    const masterRelsPath = `ppt/slideMasters/_rels/slideMaster${i + 1}.xml.rels`;

    if (useRaw && raw!.has(masterPath)) {
      entries.push({ path: masterPath, data: new TextEncoder().encode(raw!.get(masterPath)!).buffer });
    } else {
      entries.push({ path: masterPath, data: new TextEncoder().encode(parts.masters[i]).buffer });
    }

    if (useRaw && raw!.has(masterRelsPath)) {
      entries.push({ path: masterRelsPath, data: new TextEncoder().encode(raw!.get(masterRelsPath)!).buffer });
    } else {
      entries.push({ path: masterRelsPath, data: new TextEncoder().encode(parts.masterRels[i]).buffer });
    }
  }

  // 幻灯片布局
  for (let i = 0; i < parts.layouts.length; i++) {
    const layoutPath = `ppt/slideLayouts/slideLayout${i + 1}.xml`;
    const layoutRelsPath = `ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`;

    if (useRaw && raw!.has(layoutPath)) {
      entries.push({ path: layoutPath, data: new TextEncoder().encode(raw!.get(layoutPath)!).buffer });
    } else {
      entries.push({ path: layoutPath, data: new TextEncoder().encode(parts.layouts[i]).buffer });
    }

    if (useRaw && raw!.has(layoutRelsPath)) {
      entries.push({ path: layoutRelsPath, data: new TextEncoder().encode(raw!.get(layoutRelsPath)!).buffer });
    } else {
      entries.push({ path: layoutRelsPath, data: new TextEncoder().encode(buildSlideLayoutRels()).buffer });
    }
  }

  return zip(entries);
}

export function createPptx(options?: { title?: string; creator?: string; description?: string; [key: string]: any }): PptxPresentation {
  const now = new Date().toISOString();
  const theme: Theme = {
    colorScheme: {
      name: 'Office',
      colors: {
        dk1: '000000', lt1: 'FFFFFF', dk2: '000000', lt2: 'FFFFFF',
        accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000',
        accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
      },
    },
    fontScheme: { name: 'Office', majorFont: 'Calibri', minorFont: 'Calibri' },
  };
  const layout: SlideLayout = { id: '1', name: 'Blank', type: 'blank', placeholders: [] };
  const master: SlideMaster = { id: '1', layouts: [layout] };

  return {
    meta: {
      title: options?.title,
      creator: options?.creator,
      description: options?.description,
      lastModifiedBy: options?.creator,
      created: now,
      modified: now,
      ...options,
    },
    slides: [{ elements: [], layout: '1' }],
    masters: [master],
    layouts: [layout],
    theme,
    slideSize: { width: 9144000, height: 6858000 },
  };
}

export async function loadPptx(path: string): Promise<ParseResult<PptxPresentation>> {
  const { loadFromFile } = await import('../core/io.js');
  const buffer = await loadFromFile(path);
  return parsePptx(buffer);
}

export async function savePptx(pres: PptxPresentation, path: string): Promise<void> {
  const { saveToFile } = await import('../core/io.js');
  const buffer = await serializePptx(pres);
  await saveToFile(buffer, path);
}

export async function writePptxToStream(pres: PptxPresentation, stream: import('stream').Writable): Promise<void> {
  const { writeToStream } = await import('../core/io.js');
  const buffer = await serializePptx(pres);
  await writeToStream(buffer, stream);
}

function buildPackageRels(pres?: PptxPresentation): string {
  const rels: string[] = [
    `  <Relationship Id="${escapeXmlAttr('rId1')}" Type="${escapeXmlAttr('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')}" Target="${escapeXmlAttr('ppt/presentation.xml')}"/>`,
    `  <Relationship Id="${escapeXmlAttr('rId2')}" Type="${escapeXmlAttr('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties')}" Target="${escapeXmlAttr('docProps/core.xml')}"/>`,
    `  <Relationship Id="${escapeXmlAttr('rId3')}" Type="${escapeXmlAttr('http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties')}" Target="${escapeXmlAttr('docProps/app.xml')}"/>`,
  ];

  if (pres?.customProperties && pres.customProperties.length > 0) {
    rels.push(`  <Relationship Id="${escapeXmlAttr('rId4')}" Type="${escapeXmlAttr('http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties')}" Target="${escapeXmlAttr('docProps/custom.xml')}"/>`);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join('\n')}
</Relationships>`;
}

function buildAppXml(pres: PptxPresentation): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>office-meta-parser</Application>
  <Slides>${pres.slides.length}</Slides>
</Properties>`;
}

function buildPresProps(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
}

function buildSlideLayoutRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}
