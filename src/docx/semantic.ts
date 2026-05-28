import type { RawDocument } from '../core/types.js';
import type { DocxDocument } from './types.js';
import { parseMeta } from '../core/meta.js';
import { extractBody } from './parsers/body.js';
import { extractComments } from './parsers/comments.js';
import { extractStyles } from './parsers/styles.js';
import { extractRevisions } from './parsers/revision.js';
import { extractHeaders, extractFooters } from './parsers/header-footer.js';
import { extractNumbering } from './parsers/numbering.js';
import { extractFootnotes } from './parsers/footnote.js';
import { extractEndnotes } from './parsers/endnote.js';
import { extractTheme } from './parsers/theme.js';
import { extractFonts } from './parsers/font-table.js';
import { extractSettings } from './parsers/settings.js';
import { extractCommentExts } from './parsers/comments-extended.js';
import { extractPeople } from './parsers/people.js';
import { parseAppMeta } from '../core/app-meta.js';
import { parseCustomProperties } from '../core/custom-meta.js';

// 已被语义化的 XML part 路径
// 注意：endnotes.xml、fontTable.xml、settings.xml、theme1.xml、
// commentsExtended.xml、people.xml 虽有语义解析，但仍通过 extraParts 透传
const KNOWN_PARTS = new Set([
  'word/document.xml',
  'word/styles.xml',
  'word/comments.xml',
  'word/numbering.xml',
  'word/footnotes.xml',
  'docProps/core.xml',
  'docProps/app.xml',
  'docProps/custom.xml',
]);

// 已知 rels 路径
const KNOWN_RELS = new Set([
  '_rels/.rels',
  'word/_rels/document.xml.rels',
]);

function isKnownPart(path: string): boolean {
  if (KNOWN_PARTS.has(path)) return true;
  if (/^word\/header\d+\.xml$/.test(path)) return true;
  if (/^word\/footer\d+\.xml$/.test(path)) return true;
  return false;
}

export function rawToSemantic(raw: RawDocument): DocxDocument {
  const result: DocxDocument = {
    meta: raw.parts.has('docProps/core.xml') ? parseMeta(raw.parts.get('docProps/core.xml')!) : {},
    styles: extractStyles(raw),
    body: extractBody(raw),
    comments: extractComments(raw),
    trackChanges: extractRevisions(raw),
    headers: extractHeaders(raw),
    footers: extractFooters(raw),
    numbering: extractNumbering(raw),
    footnotes: extractFootnotes(raw),
  };

  // 新增语义字段
  const endnotes = extractEndnotes(raw);
  if (endnotes) result.endnotes = endnotes;

  const theme = extractTheme(raw);
  if (theme) result.theme = theme;

  const fonts = extractFonts(raw);
  if (fonts) result.fonts = fonts;

  const settings = extractSettings(raw);
  if (settings) result.settings = settings;

  const commentExts = extractCommentExts(raw);
  if (commentExts) result.commentExts = commentExts;

  const people = extractPeople(raw);
  if (people) result.people = people;

  const appXml = raw.parts.get('docProps/app.xml');
  if (appXml) result.appMeta = parseAppMeta(appXml);

  const customXml = raw.parts.get('docProps/custom.xml');
  if (customXml) result.customProperties = parseCustomProperties(customXml);

  // 存储原始 document.xml 根元素属性（命名空间、mc:Ignorable 等）
  const documentXml = raw.parts.get('word/document.xml');
  if (documentXml) {
    result.originalRootAttrs = { ...documentXml.attrs };
  }

  // 存储原始 styles.xml 根元素属性（命名空间、mc:Ignorable 等）
  const stylesXml = raw.parts.get('word/styles.xml');
  if (stylesXml) {
    result.originalStylesRootAttrs = { ...stylesXml.attrs };
  }

  // 收集未语义化的 XML part
  const extraParts = new Map<string, import('../core/types.js').ParsedNode>();
  for (const [path, node] of raw.parts) {
    if (!isKnownPart(path)) {
      extraParts.set(path, node);
    }
  }
  if (extraParts.size > 0) result.extraParts = extraParts;

  // 收集未处理的 rels
  const extraRels = new Map<string, import('../core/types.js').Relationship[]>();
  for (const [path, rels] of raw.rels) {
    if (!KNOWN_RELS.has(path)) {
      extraRels.set(path, rels);
    }
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
