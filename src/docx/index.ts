import { unzip, zip } from '../core/zip.js';
import { parseRels } from '../core/rels.js';
import { parseContentTypes } from '../core/content-type.js';
import { parseXml, serializeXml } from '../core/xml.js';
import { serializeMeta } from '../core/meta.js';
import { serializeAppMeta } from '../core/app-meta.js';
import { serializeCustomProperties } from '../core/custom-meta.js';
import { throwOnError, warnOnWarning } from '../core/validate.js';
import type { RawDocument, ParsedNode } from '../core/types.js';
import type { DocxDocument, DocxBlock } from './types.js';
import { rawToSemantic } from './semantic.js';
import { semanticToXml, serializeStyles, serializeComments, serializeHeader, serializeFooter, serializeNumbering, serializeFootnotes, serializeEndnotes, serializeTheme, serializeFontTable, serializeSettings, serializeCommentExts, serializePeople } from './serializer.js';
import { parseDocxXml } from './parser.js';
import { validateDocx } from './validate.js';

export type {
  DocxDocument,
  DocumentMeta,
  StyleDefinitions,
  ParagraphStyle,
  CharacterStyle,
  TableStyle,
  DocxBody,
  DocxBlock,
  Paragraph,
  ParagraphProperties,
  NumberingProperties,
  TextRun,
  RunProperties,
  Field,
  Table,
  TableRow,
  TableCell,
  TableProperties,
  TableRowProperties,
  TableCellProperties,
  TableBorders,
  BorderStyle,
  Image,
  Comment,
  Revision,
  NumberingDefinitions,
  AbstractNum,
  NumberingLevel,
  Num,
  Footnote,
  BookmarkStart,
  BookmarkEnd,
  ShadingStyle,
  ParagraphBorder,
  TabStop,
  FontEntry,
  DocumentSettings,
  CommentExtended,
  Person,
} from './types.js';

export type { AppMeta } from '../core/app-meta.js';
export type { CustomProperty } from '../core/custom-meta.js';

export { parseDocxXml } from './parser.js';
export { rawToSemantic } from './semantic.js';
export { semanticToXml } from './serializer.js';
export { validateDocx } from './validate.js';
export {
  updateDocxTitle, updateDocxSubject, updateDocxCreator,
  updateDocxDescription, updateDocxKeywords, updateDocxLastModifiedBy,
  updateDocxCategory, docx,
  toDocxJSON, toDocxJSONString, saveDocxJSON,
} from './meta.js';
export {
  addComment, removeComment, listComments, getCommentText,
  markCommentDone, markCommentUndone,
} from './comments.js';
export type { CommentInfo } from './comments.js';
export {
  markInsert, markDelete, addFormatChange, clearRevision,
  listRevisions, hasPendingRevisions,
  acceptAllInserts, acceptAllDeletes, rejectAllInserts, rejectAllDeletes,
} from './revisions.js';
export type { RevisionInfo } from './revisions.js';

export async function parseDocx(buffer: ArrayBuffer): Promise<{ raw: RawDocument; semantic: DocxDocument }> {
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

export async function serializeDocx(doc: DocxDocument): Promise<ArrayBuffer> {
  const issues = validateDocx(doc);
  warnOnWarning(issues);
  throwOnError(issues);

  const raw = doc.rawXmlParts;
  const useRaw = raw && raw.size > 0;

  const documentXml = useRaw && raw!.has('word/document.xml') ? raw!.get('word/document.xml')! : semanticToXml(doc);
  const metaXml = useRaw && raw!.has('docProps/core.xml') ? raw!.get('docProps/core.xml')! : serializeMeta(doc.meta);
  const stylesXml = useRaw && raw!.has('word/styles.xml') ? raw!.get('word/styles.xml')! : serializeStyles(doc.styles, doc.originalStylesRootAttrs);

  const entries: { path: string; data: ArrayBuffer }[] = [];

  // Content_Types
  if (useRaw && raw!.has('[Content_Types].xml')) {
    entries.push({ path: '[Content_Types].xml', data: new TextEncoder().encode(raw!.get('[Content_Types].xml')!).buffer });
  } else {
    entries.push({ path: '[Content_Types].xml', data: new TextEncoder().encode(buildContentTypes(doc)).buffer });
  }

  // _rels/.rels
  if (useRaw && raw!.has('_rels/.rels')) {
    entries.push({ path: '_rels/.rels', data: new TextEncoder().encode(raw!.get('_rels/.rels')!).buffer });
  } else {
    entries.push({ path: '_rels/.rels', data: new TextEncoder().encode(buildRels(doc)).buffer });
  }

  entries.push(
    { path: 'docProps/core.xml', data: new TextEncoder().encode(metaXml).buffer },
    { path: 'word/document.xml', data: new TextEncoder().encode(documentXml).buffer },
    { path: 'word/styles.xml', data: new TextEncoder().encode(stylesXml).buffer },
  );

  // word/_rels/document.xml.rels
  if (useRaw && raw!.has('word/_rels/document.xml.rels')) {
    entries.push({ path: 'word/_rels/document.xml.rels', data: new TextEncoder().encode(raw!.get('word/_rels/document.xml.rels')!).buffer });
  } else {
    entries.push({ path: 'word/_rels/document.xml.rels', data: new TextEncoder().encode(buildWordRels(doc)).buffer });
  }

  if (doc.appMeta) {
    const xml = useRaw && raw!.has('docProps/app.xml') ? raw!.get('docProps/app.xml')! : serializeAppMeta(doc.appMeta);
    entries.push({ path: 'docProps/app.xml', data: new TextEncoder().encode(xml).buffer });
  }

  if (doc.customProperties && doc.customProperties.length > 0) {
    const xml = useRaw && raw!.has('docProps/custom.xml') ? raw!.get('docProps/custom.xml')! : serializeCustomProperties(doc.customProperties);
    entries.push({ path: 'docProps/custom.xml', data: new TextEncoder().encode(xml).buffer });
  }

  // 编号定义
  if (useRaw && raw!.has('word/numbering.xml')) {
    entries.push({ path: 'word/numbering.xml', data: new TextEncoder().encode(raw!.get('word/numbering.xml')!).buffer });
  } else if (doc.numbering && (doc.numbering.abstractNums.length > 0 || doc.numbering.nums.length > 0)) {
    entries.push({ path: 'word/numbering.xml', data: new TextEncoder().encode(serializeNumbering(doc.numbering)).buffer });
  }

  // 脚注
  if (useRaw && raw!.has('word/footnotes.xml')) {
    entries.push({ path: 'word/footnotes.xml', data: new TextEncoder().encode(raw!.get('word/footnotes.xml')!).buffer });
  } else if (doc.footnotes && doc.footnotes.length > 0) {
    entries.push({ path: 'word/footnotes.xml', data: new TextEncoder().encode(serializeFootnotes(doc.footnotes)).buffer });
  }

  // 批注
  if (useRaw && raw!.has('word/comments.xml')) {
    entries.push({ path: 'word/comments.xml', data: new TextEncoder().encode(raw!.get('word/comments.xml')!).buffer });
  } else if (doc.comments && doc.comments.length > 0) {
    entries.push({ path: 'word/comments.xml', data: new TextEncoder().encode(serializeComments(doc.comments)).buffer });
  }

  // 页眉
  if (doc.headers) {
    for (let i = 0; i < doc.headers.length; i++) {
      const path = `word/header${i + 1}.xml`;
      const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializeHeader(doc.headers[i]);
      entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    }
  } else if (useRaw) {
    // 输出原始页眉（语义模型中未解析的页眉）
    for (const [path, xml] of raw!) {
      if (/^word\/header\d+\.xml$/.test(path) && !entries.some(e => e.path === path)) {
        entries.push({ path, data: new TextEncoder().encode(xml).buffer });
      }
    }
  }

  // 页脚
  if (doc.footers) {
    for (let i = 0; i < doc.footers.length; i++) {
      const path = `word/footer${i + 1}.xml`;
      const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializeFooter(doc.footers[i]);
      entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    }
  } else if (useRaw) {
    // 输出原始页脚（语义模型中未解析的页脚）
    for (const [path, xml] of raw!) {
      if (/^word\/footer\d+\.xml$/.test(path) && !entries.some(e => e.path === path)) {
        entries.push({ path, data: new TextEncoder().encode(xml).buffer });
      }
    }
  }

  // 输出有语义解析但通过 extraParts 透传的 parts
  const handledPaths = new Set<string>();

  // theme
  if (doc.theme) {
    const path = 'word/theme/theme1.xml';
    const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializeTheme(doc.theme);
    entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    handledPaths.add(path);
  }

  // fontTable
  if (doc.fonts) {
    const path = 'word/fontTable.xml';
    const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializeFontTable(doc.fonts);
    entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    handledPaths.add(path);
  }

  // settings
  if (doc.settings) {
    const path = 'word/settings.xml';
    const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializeSettings(doc.settings);
    entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    handledPaths.add(path);
  }

  // commentsExtended
  if (doc.commentExts) {
    const path = 'word/commentsExtended.xml';
    const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializeCommentExts(doc.commentExts);
    entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    handledPaths.add(path);
  }

  // people
  if (doc.people) {
    const path = 'word/people.xml';
    const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializePeople(doc.people);
    entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    handledPaths.add(path);
  }

  // endnotes
  if (doc.endnotes) {
    const path = 'word/endnotes.xml';
    const xml = useRaw && raw!.has(path) ? raw!.get(path)! : serializeEndnotes(doc.endnotes);
    entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    handledPaths.add(path);
  }

  // 输出未语义化的 extra XML parts（跳过已处理的）
  if (doc.extraParts) {
    for (const [path, node] of doc.extraParts) {
      if (handledPaths.has(path)) continue;
      entries.push({ path, data: new TextEncoder().encode(serializeXml(node)).buffer });
    }
  }

  // 输出未处理的 extra rels（跳过已生成的 rels 路径）
  const generatedRels = new Set(['_rels/.rels', 'word/_rels/document.xml.rels']);
  if (doc.extraRels) {
    for (const [path, rels] of doc.extraRels) {
      if (generatedRels.has(path)) continue;
      const relsXml = rels.map(rel => {
        let attrs = `Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"`;
        if (rel.targetMode) attrs += ` TargetMode="${rel.targetMode}"`;
        return `  <Relationship ${attrs}/>`;
      }).join('\n');
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${relsXml}\n</Relationships>`;
      entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    }
  }

  // 输出二进制 entries（图片等）
  if (doc.extraEntries) {
    for (const entry of doc.extraEntries) {
      entries.push({ path: entry.path, data: entry.data });
    }
  }

  return zip(entries);
}

export function serializeDocxXml(node: ParsedNode): string {
  return serializeXml(node);
}

export function createDocx(options?: { title?: string; creator?: string; description?: string; [key: string]: any }): DocxDocument {
  const now = new Date().toISOString();
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
    styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
    body: { blocks: [] },
  };
}

export async function loadDocx(path: string): Promise<{ raw: RawDocument; semantic: DocxDocument }> {
  const { loadFromFile } = await import('../core/io.js');
  const buffer = await loadFromFile(path);
  return parseDocx(buffer);
}

export async function saveDocx(doc: DocxDocument, path: string): Promise<void> {
  const { saveToFile } = await import('../core/io.js');
  const buffer = await serializeDocx(doc);
  await saveToFile(buffer, path);
}

export async function writeDocxToStream(doc: DocxDocument, stream: import('stream').Writable): Promise<void> {
  const { writeToStream } = await import('../core/io.js');
  const buffer = await serializeDocx(doc);
  await writeToStream(buffer, stream);
}

function buildContentTypes(doc: DocxDocument): string {
  const overrides: string[] = [
    `  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
    `  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
    `  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
  ];

  if (doc.comments && doc.comments.length > 0) {
    overrides.push(`  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`);
  }

  if (doc.headers) {
    for (let i = 0; i < doc.headers.length; i++) {
      overrides.push(`  <Override PartName="/word/header${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`);
    }
  }

  if (doc.footers) {
    for (let i = 0; i < doc.footers.length; i++) {
      overrides.push(`  <Override PartName="/word/footer${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`);
    }
  }

  if (doc.numbering && (doc.numbering.abstractNums.length > 0 || doc.numbering.nums.length > 0)) {
    overrides.push(`  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`);
  }

  if (doc.footnotes && doc.footnotes.length > 0) {
    overrides.push(`  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`);
  }

  if (doc.appMeta) {
    overrides.push(`  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`);
  }

  if (doc.customProperties && doc.customProperties.length > 0) {
    overrides.push(`  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`);
  }

  // 合并 extra content types（以生成的为准，补充缺失的 Override）
  if (doc.extraContentTypes) {
    const generatedPartNames = new Set<string>();
    for (const line of overrides) {
      const m = line.match(/PartName="([^"]+)"/);
      if (m) generatedPartNames.add(m[1]);
    }
    for (const ct of doc.extraContentTypes) {
      if (ct.partName && !generatedPartNames.has(ct.partName)) {
        overrides.push(`  <Override PartName="${ct.partName}" ContentType="${ct.contentType}"/>`);
        generatedPartNames.add(ct.partName);
      }
    }
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${overrides.join('\n')}
</Types>`;
}

function buildRels(doc?: DocxDocument): string {
  const rels: string[] = [
    `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`,
    `  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
  ];

  if (doc?.appMeta) {
    rels.push(`  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`);
  }

  if (doc?.customProperties && doc.customProperties.length > 0) {
    const id = doc.appMeta ? 'rId4' : 'rId3';
    rels.push(`  <Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>`);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join('\n')}
</Relationships>`;
}

function buildWordRels(doc: DocxDocument): string {
  const rels: { id: string; type: string; target: string; targetMode?: string }[] = [];

  // 添加样式关系
  rels.push({
    id: 'rId1',
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
    target: 'styles.xml',
  });

  // 收集所有超链接关系
  const hyperlinks = collectHyperlinks(doc);
  for (const hyperlink of hyperlinks) {
    if (hyperlink.url) {
      rels.push({
        id: hyperlink.relationshipId,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        target: hyperlink.url,
        targetMode: 'External',
      });
    }
  }

  // 收集所有图片关系
  const images = collectImages(doc);
  for (const image of images) {
    rels.push({
      id: image.relationshipId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: `media/${image.relationshipId}.png`,
    });
  }

  // 添加批注关系
  if (doc.comments && doc.comments.length > 0) {
    rels.push({
      id: 'rId100',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
      target: 'comments.xml',
    });
  }

  // 添加页眉页脚关系
  if (doc.headers) {
    for (let i = 0; i < doc.headers.length; i++) {
      rels.push({
        id: doc.headers[i].id,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
        target: `header${i + 1}.xml`,
      });
    }
  }

  if (doc.footers) {
    for (let i = 0; i < doc.footers.length; i++) {
      rels.push({
        id: doc.footers[i].id,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
        target: `footer${i + 1}.xml`,
      });
    }
  }

  const relsXml = rels.map(rel => {
    let attrs = `Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"`;
    if (rel.targetMode) {
      attrs += ` TargetMode="${rel.targetMode}"`;
    }
    return `  <Relationship ${attrs}/>`;
  }).join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relsXml}
</Relationships>`;
}

function collectHyperlinks(doc: DocxDocument): { relationshipId: string; url?: string }[] {
  const hyperlinks: { relationshipId: string; url?: string }[] = [];

  function collectFromBlocks(blocks: DocxBlock[]) {
    for (const block of blocks) {
      if (block.type === 'hyperlink') {
        hyperlinks.push({
          relationshipId: block.relationshipId,
          url: block.url,
        });
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            collectFromBlocks(cell.blocks);
          }
        }
      }
    }
  }

  collectFromBlocks(doc.body.blocks);
  return hyperlinks;
}

function collectImages(doc: DocxDocument): { relationshipId: string }[] {
  const images: { relationshipId: string }[] = [];

  function collectFromBlocks(blocks: DocxBlock[]) {
    for (const block of blocks) {
      if (block.type === 'image') {
        images.push({
          relationshipId: block.relationshipId,
        });
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            collectFromBlocks(cell.blocks);
          }
        }
      }
    }
  }

  collectFromBlocks(doc.body.blocks);
  return images;
}
