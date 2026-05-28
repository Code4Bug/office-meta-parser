import type { RawDocument, ParsedNode } from '../../core/types.js';
import type {
  Theme, ColorScheme, FontScheme, FormatScheme,
  FillStyle, BorderStyle, EffectStyle, ShadowStyle, GlowStyle, SoftEdgeStyle,
} from '../types.js';
import { findChild, findChildAny } from './utils.js';
import { parseColorValue, parseGradientFill } from './style.js';

export function extractTheme(raw: RawDocument): Theme | undefined {
  for (const [path, xml] of raw.parts) {
    if (path.startsWith('ppt/theme/theme') && path.endsWith('.xml')) {
      return parseTheme(xml);
    }
  }
  return undefined;
}

function parseTheme(node: ParsedNode): Theme {
  const themeElements = findChild(node, 'a:themeElements');
  if (!themeElements) {
    return { colorScheme: { colors: {} }, fontScheme: {} };
  }

  const theme: Theme = {
    colorScheme: parseColorScheme(themeElements),
    fontScheme: parseFontScheme(themeElements),
  };

  // Parse fmtScheme from themeElements or directly
  const fmtScheme = findChild(themeElements, 'a:fmtScheme');
  if (fmtScheme) {
    theme.formatScheme = parseFormatScheme(fmtScheme);
  }

  return theme;
}

function parseColorScheme(themeElements: ParsedNode): ColorScheme {
  const clrScheme = findChild(themeElements, 'a:clrScheme');
  if (!clrScheme) return { colors: {} };

  const name = clrScheme.attrs['name'];
  const colors: Record<string, string> = {};

  for (const child of clrScheme.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:dk1' || child.tag === 'a:lt1' || child.tag === 'a:dk2' || child.tag === 'a:lt2' ||
        child.tag === 'a:accent1' || child.tag === 'a:accent2' || child.tag === 'a:accent3' ||
        child.tag === 'a:accent4' || child.tag === 'a:accent5' || child.tag === 'a:accent6' ||
        child.tag === 'a:hlink' || child.tag === 'a:folHlink') {
      const color = parseColorValue(child);
      if (color) colors[child.tag.replace('a:', '')] = color;
    }
  }

  const scheme: ColorScheme = { colors };
  if (name) scheme.name = name;
  return scheme;
}

function parseFontScheme(themeElements: ParsedNode): FontScheme {
  const fontScheme = findChild(themeElements, 'a:fontScheme');
  if (!fontScheme) return {};

  const name = fontScheme.attrs['name'];
  const majorFont = findChild(fontScheme, 'a:majorFont');
  const minorFont = findChild(fontScheme, 'a:minorFont');

  const scheme: FontScheme = {};
  if (name) scheme.name = name;
  if (majorFont) {
    const latin = findChild(majorFont, 'a:latin');
    if (latin?.attrs['typeface']) scheme.majorFont = latin.attrs['typeface'];
  }
  if (minorFont) {
    const latin = findChild(minorFont, 'a:latin');
    if (latin?.attrs['typeface']) scheme.minorFont = latin.attrs['typeface'];
  }

  return scheme;
}

function parseFormatScheme(node: ParsedNode): FormatScheme {
  const fs: FormatScheme = {};

  // Fill styles
  const fillStyleLst = findChild(node, 'a:fillStyleLst');
  if (fillStyleLst) {
    fs.fillStyles = parseFillStyleList(fillStyleLst);
  }

  // Line styles
  const lnStyleLst = findChild(node, 'a:lnStyleLst');
  if (lnStyleLst) {
    fs.lineStyles = parseLineStyleList(lnStyleLst);
  }

  // Effect styles
  const effectStyleLst = findChild(node, 'a:effectStyleLst');
  if (effectStyleLst) {
    fs.effectStyles = parseEffectStyleList(effectStyleLst);
  }

  // Background fill styles
  const bgFillStyleLst = findChild(node, 'a:bgFillStyleLst');
  if (bgFillStyleLst) {
    fs.bgFillStyles = parseFillStyleList(bgFillStyleLst);
  }

  return fs;
}

function parseFillStyleList(node: ParsedNode): FillStyle[] {
  const fills: FillStyle[] = [];
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:solidFill') {
      const color = parseColorValue(child);
      if (color) fills.push({ type: 'solid', color });
    } else if (child.tag === 'a:gradFill') {
      fills.push({ type: 'gradient', gradientFill: parseGradientFill(child) });
    } else if (child.tag === 'a:pattFill') {
      const pf: { preset?: string; fgColor?: string; bgColor?: string } = {};
      if (child.attrs['prst']) pf.preset = child.attrs['prst'];
      const fgClr = findChild(child, 'a:fgClr');
      if (fgClr) pf.fgColor = parseColorValue(fgClr);
      const bgClr = findChild(child, 'a:bgClr');
      if (bgClr) pf.bgColor = parseColorValue(bgClr);
      fills.push({ type: 'pattern', patternFill: pf });
    } else if (child.tag === 'a:noFill') {
      fills.push({ type: 'none' });
    }
  }
  return fills;
}

function parseLineStyleList(node: ParsedNode): BorderStyle[] {
  const lines: BorderStyle[] = [];
  for (const child of node.children) {
    if (typeof child === 'string' || child.tag !== 'a:ln') continue;
    const border: BorderStyle = {};
    if (child.attrs['w']) border.width = parseInt(child.attrs['w'], 10) / 12700;
    const solidFill = findChild(child, 'a:solidFill');
    if (solidFill) border.color = parseColorValue(solidFill);
    if (child.attrs['cmpd']) border.compound = child.attrs['cmpd'] as BorderStyle['compound'];
    if (child.attrs['cap']) border.cap = child.attrs['cap'] as BorderStyle['cap'];
    if (Object.keys(border).length > 0) lines.push(border);
  }
  return lines;
}

function parseEffectStyleList(node: ParsedNode): EffectStyle[] {
  const effects: EffectStyle[] = [];
  for (const child of node.children) {
    if (typeof child === 'string' || child.tag !== 'a:effectStyle') continue;
    const effectLst = findChild(child, 'a:effectLst');
    if (!effectLst) { effects.push({}); continue; }

    const effect: EffectStyle = {};
    const outerShdw = findChild(effectLst, 'a:outerShdw');
    if (outerShdw) {
      effect.shadow = { type: 'outer' };
      if (outerShdw.attrs['blurRad']) effect.shadow.blur = parseInt(outerShdw.attrs['blurRad'], 10) / 12700;
      const color = parseColorValue(outerShdw);
      if (color) effect.shadow.color = color;
    }
    const glow = findChild(effectLst, 'a:glow');
    if (glow) {
      effect.glow = {};
      if (glow.attrs['rad']) effect.glow.radius = parseInt(glow.attrs['rad'], 10) / 12700;
      const color = parseColorValue(glow);
      if (color) effect.glow.color = color;
    }
    const softEdge = findChild(effectLst, 'a:softEdge');
    if (softEdge) {
      effect.softEdge = {};
      if (softEdge.attrs['rad']) effect.softEdge.radius = parseInt(softEdge.attrs['rad'], 10) / 12700;
    }
    effects.push(effect);
  }
  return effects;
}
