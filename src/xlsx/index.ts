import { unzip, zip } from '../core/zip.js';
import { parseRels } from '../core/rels.js';
import { parseContentTypes } from '../core/content-type.js';
import { parseXml, serializeXml, escapeXmlAttr } from '../core/xml.js';
import { serializeMeta } from '../core/meta.js';
import { serializeAppMeta } from '../core/app-meta.js';
import { serializeCustomProperties } from '../core/custom-meta.js';
import { throwOnError, warnOnWarning } from '../core/validate.js';
import type { RawDocument, ParsedNode } from '../core/types.js';
import type { XlsxWorkbook } from './types.js';
import { rawToSemantic } from './semantic.js';
import { semanticToXml } from './serializer.js';
import { validateXlsx } from './validate.js';

export type {
  XlsxWorkbook,
  Sheet,
  Cell,
  MergedCell,
  CellStyleDefinitions,
  CellStyle,
  FontDefinition,
  FillDefinition,
  BorderDefinition,
  BorderStyle,
  Alignment,
  NumberFormatDefinition,
  FrozenPanes,
  SheetSelection,
  OutlineGroup,
  PageSetup,
  SheetHeaderFooter,
  SheetComment,
  ExcelTable,
  TableColumn,
  DefinedName,
  ThemeDefinition,
  RichTextRun,
  GradientFill,
  GradientStop,
} from './types.js';

export type { AppMeta } from '../core/app-meta.js';
export type { CustomProperty } from '../core/custom-meta.js';

export { parseXlsxXml } from './parser.js';
export { rawToSemantic } from './semantic.js';
export { semanticToXml } from './serializer.js';
export { validateXlsx } from './validate.js';
export {
  updateXlsxTitle, updateXlsxSubject, updateXlsxCreator,
  updateXlsxDescription, updateXlsxKeywords, updateXlsxLastModifiedBy,
  updateXlsxCategory, xlsx,
  toXlsxJSON, toXlsxJSONString, saveXlsxJSON,
} from './meta.js';
export {
  addComment, removeComment, listComments, listSheetComments, getCommentText, updateComment,
} from './comments.js';
export type { CommentInfo } from './comments.js';

export async function parseXlsx(buffer: ArrayBuffer): Promise<{ raw: RawDocument; semantic: XlsxWorkbook }> {
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

export async function serializeXlsx(wb: XlsxWorkbook): Promise<ArrayBuffer> {
  const issues = validateXlsx(wb);
  warnOnWarning(issues);
  throwOnError(issues);

  const parts = semanticToXml(wb);
  const raw = wb.rawXmlParts;
  const useRaw = raw && raw.size > 0;

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
    entries.push({ path: '_rels/.rels', data: new TextEncoder().encode(buildPackageRels(wb)).buffer });
  }

  const metaXml = useRaw && raw!.has('docProps/core.xml') ? raw!.get('docProps/core.xml')! : serializeMeta(wb.meta);
  entries.push({ path: 'docProps/core.xml', data: new TextEncoder().encode(metaXml).buffer });

  // xl/workbook.xml
  if (useRaw && raw!.has('xl/workbook.xml')) {
    entries.push({ path: 'xl/workbook.xml', data: new TextEncoder().encode(raw!.get('xl/workbook.xml')!).buffer });
  } else {
    entries.push({ path: 'xl/workbook.xml', data: new TextEncoder().encode(parts.workbook).buffer });
  }

  // xl/_rels/workbook.xml.rels
  if (useRaw && raw!.has('xl/_rels/workbook.xml.rels')) {
    entries.push({ path: 'xl/_rels/workbook.xml.rels', data: new TextEncoder().encode(raw!.get('xl/_rels/workbook.xml.rels')!).buffer });
  } else {
    entries.push({ path: 'xl/_rels/workbook.xml.rels', data: new TextEncoder().encode(parts.rels).buffer });
  }

  // xl/sharedStrings.xml
  if (useRaw && raw!.has('xl/sharedStrings.xml')) {
    entries.push({ path: 'xl/sharedStrings.xml', data: new TextEncoder().encode(raw!.get('xl/sharedStrings.xml')!).buffer });
  } else {
    entries.push({ path: 'xl/sharedStrings.xml', data: new TextEncoder().encode(parts.sharedStrings).buffer });
  }

  // worksheets
  for (let i = 0; i < parts.worksheets.length; i++) {
    const path = `xl/worksheets/sheet${i + 1}.xml`;
    if (useRaw && raw!.has(path)) {
      entries.push({ path, data: new TextEncoder().encode(raw!.get(path)!).buffer });
    } else {
      entries.push({ path, data: new TextEncoder().encode(parts.worksheets[i]).buffer });
    }
  }

  // sheet rels（仅在源文件有或工作表有实际关系时输出）
  for (let i = 0; i < parts.sheetRels.length; i++) {
    const path = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
    if (useRaw && raw!.has(path)) {
      entries.push({ path, data: new TextEncoder().encode(raw!.get(path)!).buffer });
    } else {
      const sheet = wb.sheets[i];
      const hasRels = (sheet.tables && sheet.tables.length > 0) ||
                      (sheet.comments && sheet.comments.length > 0) ||
                      (sheet.images && sheet.images.length > 0);
      if (hasRels) {
        entries.push({ path, data: new TextEncoder().encode(parts.sheetRels[i]).buffer });
      }
    }
  }

  // comments
  for (const c of parts.comments) {
    if (useRaw && raw!.has(c.path)) {
      entries.push({ path: c.path, data: new TextEncoder().encode(raw!.get(c.path)!).buffer });
    } else {
      entries.push({ path: c.path, data: new TextEncoder().encode(c.xml).buffer });
    }
  }

  if (wb.appMeta) {
    const xml = useRaw && raw!.has('docProps/app.xml') ? raw!.get('docProps/app.xml')! : serializeAppMeta(wb.appMeta);
    entries.push({ path: 'docProps/app.xml', data: new TextEncoder().encode(xml).buffer });
  }

  if (wb.customProperties && wb.customProperties.length > 0) {
    const xml = useRaw && raw!.has('docProps/custom.xml') ? raw!.get('docProps/custom.xml')! : serializeCustomProperties(wb.customProperties);
    entries.push({ path: 'docProps/custom.xml', data: new TextEncoder().encode(xml).buffer });
  }

  // 输出未语义化的 extra XML parts
  if (wb.extraParts) {
    for (const [path, node] of wb.extraParts) {
      entries.push({ path, data: new TextEncoder().encode(serializeXml(node)).buffer });
    }
  }

  // 输出未处理的 extra rels（跳过已生成的 rels 路径）
  const generatedRels = new Set(['_rels/.rels', 'xl/_rels/workbook.xml.rels']);
  if (wb.extraRels) {
    for (const [path, rels] of wb.extraRels) {
      if (generatedRels.has(path)) continue;
      const relsXml = rels.map(rel => {
        let attrs = `Id="${escapeXmlAttr(rel.id)}" Type="${escapeXmlAttr(rel.type)}" Target="${escapeXmlAttr(rel.target)}"`;
        if (rel.targetMode) attrs += ` TargetMode="${escapeXmlAttr(rel.targetMode)}"`;
        return `  <Relationship ${attrs}/>`;
      }).join('\n');
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${relsXml}\n</Relationships>`;
      entries.push({ path, data: new TextEncoder().encode(xml).buffer });
    }
  }

  // 输出二进制 entries（图片等）
  if (wb.extraEntries) {
    for (const entry of wb.extraEntries) {
      entries.push({ path: entry.path, data: entry.data });
    }
  }

  return zip(entries);
}

export function createXlsx(options?: { title?: string; creator?: string; sheetName?: string; [key: string]: any }): XlsxWorkbook {
  const now = new Date().toISOString();
  return {
    meta: {
      title: options?.title,
      creator: options?.creator,
      lastModifiedBy: options?.creator,
      created: now,
      modified: now,
      ...options,
    },
    sheets: [{
      name: options?.sheetName ?? 'Sheet1',
      cells: [],
      mergedCells: [],
      columnWidths: [],
      rowHeights: [],
      hyperlinks: [],
    }],
    styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
    sharedStrings: [],
  };
}

export async function loadXlsx(path: string): Promise<{ raw: RawDocument; semantic: XlsxWorkbook }> {
  const { loadFromFile } = await import('../core/io.js');
  const buffer = await loadFromFile(path);
  return parseXlsx(buffer);
}

export async function saveXlsx(wb: XlsxWorkbook, path: string): Promise<void> {
  const { saveToFile } = await import('../core/io.js');
  const buffer = await serializeXlsx(wb);
  await saveToFile(buffer, path);
}

export async function writeXlsxToStream(wb: XlsxWorkbook, stream: import('stream').Writable): Promise<void> {
  const { writeToStream } = await import('../core/io.js');
  const buffer = await serializeXlsx(wb);
  await writeToStream(buffer, stream);
}

function buildPackageRels(wb?: XlsxWorkbook): string {
  const rels: string[] = [
    `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`,
    `  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
  ];

  if (wb?.appMeta) {
    rels.push(`  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`);
  }

  if (wb?.customProperties && wb.customProperties.length > 0) {
    const id = wb.appMeta ? 'rId4' : 'rId3';
    rels.push(`  <Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>`);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join('\n')}
</Relationships>`;
}
