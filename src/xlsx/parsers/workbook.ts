import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { DefinedName, ThemeDefinition } from '../types.js';
import { findChild } from './utils.js';

export function extractDefinedNames(raw: RawDocument): DefinedName[] | undefined {
  const workbookXml = raw.parts.get('xl/workbook.xml');
  if (!workbookXml) return undefined;

  const definedNames = findChild(workbookXml, 'definedNames');
  if (!definedNames) return undefined;

  const names: DefinedName[] = [];
  for (const child of definedNames.children) {
    if (typeof child === 'string' || child.tag !== 'definedName') continue;
    const name = child.attrs['name'];
    if (!name) continue;

    const formula = getDirectText(child);
    const dn: DefinedName = { name, formula };
    if (child.attrs['localSheetId']) dn.localSheetId = parseInt(child.attrs['localSheetId'], 10);
    if (child.attrs['hidden'] === '1') dn.hidden = true;
    names.push(dn);
  }

  return names.length > 0 ? names : undefined;
}

export function extractTheme(raw: RawDocument): ThemeDefinition | undefined {
  const themeXml = raw.parts.get('xl/theme/theme1.xml');
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
            // Strip namespace prefix for the key
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

/** findChild that matches both 'tag' and 'ns:tag' */
function findChildAny(node: ParsedNode, tag: string): ParsedNode | undefined {
  for (const child of node.children) {
    if (typeof child !== 'string') {
      if (child.tag === tag || child.tag.endsWith(':' + tag)) return child;
    }
  }
  return undefined;
}

function getDirectText(node: ParsedNode): string {
  for (const child of node.children) {
    if (typeof child === 'string') return child;
  }
  return '';
}
