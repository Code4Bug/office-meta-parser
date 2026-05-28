import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { ThemeDefinition } from '../../xlsx/types.js';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

export function serializeTheme(theme: ThemeDefinition): string {
  const colorScheme = theme.colorScheme || {};
  const colorChildren: ParsedNode[] = [];

  const colorMap: Record<string, string> = {
    dk1: 'dk1', lt1: 'lt1', dk2: 'dk2', lt2: 'lt2',
    accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
    accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
    hlink: 'hlink', folHlink: 'folHlink',
  };

  for (const [key, tag] of Object.entries(colorMap)) {
    const val = colorScheme[key];
    if (val) {
      colorChildren.push({
        tag: `a:${tag}`,
        attrs: {},
        children: [{ tag: 'a:srgbClr', attrs: { val }, children: [] }],
      });
    }
  }

  const fontSchemeChildren: ParsedNode[] = [];
  if (theme.majorFont) {
    fontSchemeChildren.push({
      tag: 'a:majorFont',
      attrs: {},
      children: [{ tag: 'a:latin', attrs: { typeface: theme.majorFont }, children: [] }],
    });
  }
  if (theme.minorFont) {
    fontSchemeChildren.push({
      tag: 'a:minorFont',
      attrs: {},
      children: [{ tag: 'a:latin', attrs: { typeface: theme.minorFont }, children: [] }],
    });
  }

  const themeElements: ParsedNode[] = [];
  if (colorChildren.length > 0) {
    themeElements.push({ tag: 'a:clrScheme', attrs: { name: '' }, children: colorChildren });
  }
  if (fontSchemeChildren.length > 0) {
    themeElements.push({ tag: 'a:fontScheme', attrs: { name: '' }, children: fontSchemeChildren });
  }

  const root: ParsedNode = {
    tag: 'a:theme',
    attrs: { 'xmlns:a': A_NS, name: 'Office' },
    children: themeElements.length > 0
      ? [{ tag: 'a:themeElements', attrs: {}, children: themeElements }]
      : [],
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
