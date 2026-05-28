import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { ThemeDefinition } from '../../xlsx/types.js';

export function extractTheme(raw: RawDocument): ThemeDefinition | undefined {
  const themeXml = raw.parts.get('word/theme/theme1.xml');
  if (!themeXml) return undefined;

  const theme: ThemeDefinition = {};
  const themeElements = findChildAny(themeXml, 'themeElements');
  if (!themeElements) return undefined;

  const clrScheme = findChildAny(themeElements, 'clrScheme');
  if (clrScheme) {
    theme.colorScheme = {};
    for (const child of clrScheme.children) {
      if (typeof child === 'string') continue;
      if (child.children.length > 0) {
        const colorNode = child.children.find(c => typeof c !== 'string') as ParsedNode | undefined;
        if (colorNode) {
          const val = colorNode.attrs['lastClr'] || colorNode.attrs['val'];
          if (val) {
            const key = child.tag.includes(':') ? child.tag.split(':')[1] : child.tag;
            theme.colorScheme[key] = val;
          }
        }
      }
    }
  }

  const fontScheme = findChildAny(themeElements, 'fontScheme');
  if (fontScheme) {
    const majorFont = findChildAny(fontScheme, 'majorFont');
    if (majorFont) {
      const latin = findChildAny(majorFont, 'latin');
      if (latin) theme.majorFont = latin.attrs['typeface'];
    }
    const minorFont = findChildAny(fontScheme, 'minorFont');
    if (minorFont) {
      const latin = findChildAny(minorFont, 'latin');
      if (latin) theme.minorFont = latin.attrs['typeface'];
    }
  }

  return theme;
}

function findChildAny(node: ParsedNode, tag: string): ParsedNode | undefined {
  for (const child of node.children) {
    if (typeof child !== 'string') {
      if (child.tag === tag || child.tag.endsWith(':' + tag)) return child;
    }
  }
  return undefined;
}
