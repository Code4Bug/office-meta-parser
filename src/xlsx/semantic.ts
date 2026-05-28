import type { RawDocument } from '../core/types.js';
import type { XlsxWorkbook } from './types.js';
import { parseMeta } from '../core/meta.js';
import { extractSheets } from './parsers/sheet.js';
import { extractSharedStrings } from './parsers/strings.js';
import { extractStyles } from './parsers/styles.js';
import { extractDefinedNames, extractTheme } from './parsers/workbook.js';
import { parseAppMeta } from '../core/app-meta.js';
import { parseCustomProperties } from '../core/custom-meta.js';

// 已被语义化的 XML part 路径
// 注意：xl/styles.xml 和 xl/theme/theme1.xml 不在此列，通过 extraParts 透传
const KNOWN_PARTS = new Set([
  'xl/workbook.xml',
  'xl/sharedStrings.xml',
  'docProps/core.xml',
  'docProps/app.xml',
  'docProps/custom.xml',
]);

function isKnownPart(path: string): boolean {
  if (KNOWN_PARTS.has(path)) return true;
  if (/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) return true;
  if (/^xl\/tables\/table\d+\.xml$/.test(path)) return true;
  if (/^xl\/pivotTables\/.+\.xml$/.test(path)) return true;
  if (/^xl\/pivotCache\/.+\.xml$/.test(path)) return true;
  return false;
}

export function rawToSemantic(raw: RawDocument): XlsxWorkbook {
  const { sheets, authors } = extractSheets(raw);
  const definedNames = extractDefinedNames(raw);

  if (definedNames) {
    for (const dn of definedNames) {
      if (dn.name === '_xlnm.Print_Titles' && dn.localSheetId !== undefined && sheets[dn.localSheetId]) {
        sheets[dn.localSheetId].printTitles = dn.formula;
      }
    }
  }

  const result: XlsxWorkbook = {
    meta: raw.parts.has('docProps/core.xml') ? parseMeta(raw.parts.get('docProps/core.xml')!) : {},
    sheets,
    styles: extractStyles(raw),
    sharedStrings: extractSharedStrings(raw),
    definedNames,
    theme: extractTheme(raw),
  };

  if (authors.length > 0) result.authors = authors;

  const appXml = raw.parts.get('docProps/app.xml');
  if (appXml) result.appMeta = parseAppMeta(appXml);

  const customXml = raw.parts.get('docProps/custom.xml');
  if (customXml) result.customProperties = parseCustomProperties(customXml);

  // 收集未语义化的 XML part
  const extraParts = new Map<string, import('../core/types.js').ParsedNode>();
  for (const [path, node] of raw.parts) {
    if (!isKnownPart(path)) {
      extraParts.set(path, node);
    }
  }
  if (extraParts.size > 0) result.extraParts = extraParts;

  // 收集未处理的 rels（排除由 serializer 生成的）
  const extraRels = new Map<string, import('../core/types.js').Relationship[]>();
  for (const [path, rels] of raw.rels) {
    extraRels.set(path, rels);
  }
  if (extraRels.size > 0) result.extraRels = extraRels;

  // 存储原始 content types
  if (raw.contentTypes.length > 0) {
    result.extraContentTypes = raw.contentTypes;
  }

  // 收集非 XML 的 entry（图片等二进制文件）
  const extraEntries = raw.entries.filter(e =>
    !e.path.endsWith('.xml') && !e.path.endsWith('.rels') && e.path !== '[Content_Types].xml'
  );
  if (extraEntries.length > 0) result.extraEntries = extraEntries;

  // 存储所有 XML 和 rels 的原始文本（用于 round-trip 还原）
  const rawXmlParts = new Map<string, string>();
  for (const entry of raw.entries) {
    if (entry.path.endsWith('.xml') || entry.path.endsWith('.rels')) {
      rawXmlParts.set(entry.path, new TextDecoder().decode(entry.data));
    }
  }
  if (rawXmlParts.size > 0) result.rawXmlParts = rawXmlParts;

  return result;
}
