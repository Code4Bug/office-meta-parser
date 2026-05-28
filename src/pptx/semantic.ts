import type { RawDocument } from '../core/types.js';
import type { PptxPresentation } from './types.js';
import { parseMeta } from '../core/meta.js';
import { extractSlides, extractSlideSize, extractNotesSize } from './parsers/slide.js';
import { extractMasters, extractLayouts } from './parsers/master.js';
import { extractTheme } from './parsers/theme.js';
import { parseAppMeta } from '../core/app-meta.js';
import { parseCustomProperties } from '../core/custom-meta.js';

export function rawToSemantic(raw: RawDocument): PptxPresentation {
  const result: PptxPresentation = {
    meta: raw.parts.has('docProps/core.xml') ? parseMeta(raw.parts.get('docProps/core.xml')!) : {},
    slides: extractSlides(raw),
    masters: extractMasters(raw),
    layouts: extractLayouts(raw),
  };

  const theme = extractTheme(raw);
  if (theme) result.theme = theme;

  const slideSize = extractSlideSize(raw);
  if (slideSize) result.slideSize = slideSize;

  const notesSize = extractNotesSize(raw);
  if (notesSize) result.notesSize = notesSize;

  const appXml = raw.parts.get('docProps/app.xml');
  if (appXml) result.appMeta = parseAppMeta(appXml);

  const customXml = raw.parts.get('docProps/custom.xml');
  if (customXml) result.customProperties = parseCustomProperties(customXml);

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
