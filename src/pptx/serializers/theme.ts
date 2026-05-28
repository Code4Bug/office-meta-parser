import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { Theme, FormatScheme, FillStyle, BorderStyle, EffectStyle } from '../types.js';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

export function serializeTheme(theme: Theme): string {
  const themeElements: ParsedNode[] = [
    serializeColorScheme(theme.colorScheme),
    serializeFontScheme(theme.fontScheme),
    serializeFormatScheme(theme.formatScheme),
  ];

  const root: ParsedNode = {
    tag: 'a:theme',
    attrs: { 'xmlns:a': A_NS, name: 'Office Theme' },
    children: [
      {
        tag: 'a:themeElements',
        attrs: {} as Record<string, string>,
        children: themeElements,
      },
    ],
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeColorScheme(scheme: Theme['colorScheme']): ParsedNode {
  const children: ParsedNode[] = [];

  const colorEntries: [string, string][] = [
    ['dk1', 'FFFFFF'],
    ['lt1', '000000'],
    ['dk2', '000000'],
    ['lt2', 'FFFFFF'],
    ['accent1', '4472C4'],
    ['accent2', 'ED7D31'],
    ['accent3', 'A5A5A5'],
    ['accent4', 'FFC000'],
    ['accent5', '5B9BD5'],
    ['accent6', '70AD47'],
    ['hlink', '0563C1'],
    ['folHlink', '954F72'],
  ];

  for (const [name, defaultColor] of colorEntries) {
    const color = scheme.colors[name] || defaultColor;
    children.push({
      tag: `a:${name}`,
      attrs: {} as Record<string, string>,
      children: [
        { tag: 'a:srgbClr', attrs: { val: color }, children: [] },
      ],
    });
  }

  const result: ParsedNode = {
    tag: 'a:clrScheme',
    attrs: { name: scheme.name || 'Office' },
    children,
  };
  return result;
}

function serializeFontScheme(scheme: Theme['fontScheme']): ParsedNode {
  const majorFont = scheme.majorFont || 'Calibri';
  const minorFont = scheme.minorFont || 'Calibri';

  return {
    tag: 'a:fontScheme',
    attrs: { name: scheme.name || 'Office' },
    children: [
      {
        tag: 'a:majorFont',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'a:latin', attrs: { typeface: majorFont }, children: [] },
          { tag: 'a:ea', attrs: { typeface: '' }, children: [] },
          { tag: 'a:cs', attrs: { typeface: '' }, children: [] },
        ],
      },
      {
        tag: 'a:minorFont',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'a:latin', attrs: { typeface: minorFont }, children: [] },
          { tag: 'a:ea', attrs: { typeface: '' }, children: [] },
          { tag: 'a:cs', attrs: { typeface: '' }, children: [] },
        ],
      },
    ],
  };
}

function serializeFormatScheme(fmtScheme?: FormatScheme): ParsedNode {
  const phClr: ParsedNode = { tag: 'a:schemeClr', attrs: { val: 'phClr' }, children: [] };
  const defaultSolidFill: ParsedNode = { tag: 'a:solidFill', attrs: {} as Record<string, string>, children: [phClr] };
  const defaultLn: ParsedNode = {
    tag: 'a:ln',
    attrs: { w: '9525' },
    children: [{ tag: 'a:solidFill', attrs: {} as Record<string, string>, children: [phClr] }],
  };
  const defaultEffectStyle: ParsedNode = {
    tag: 'a:effectStyle',
    attrs: {} as Record<string, string>,
    children: [{ tag: 'a:effectLst', attrs: {} as Record<string, string>, children: [] }],
  };

  // Serialize fill styles
  const fillStyles: ParsedNode[] = fmtScheme?.fillStyles
    ? fmtScheme.fillStyles.map(serializeThemeFill)
    : [defaultSolidFill, defaultSolidFill, defaultSolidFill];

  const lnStyles: ParsedNode[] = fmtScheme?.lineStyles
    ? fmtScheme.lineStyles.map(serializeThemeLine)
    : [defaultLn, defaultLn, defaultLn];

  const effectStyles: ParsedNode[] = fmtScheme?.effectStyles
    ? fmtScheme.effectStyles.map(serializeThemeEffect)
    : [defaultEffectStyle, defaultEffectStyle, defaultEffectStyle];

  const bgFillStyles: ParsedNode[] = fmtScheme?.bgFillStyles
    ? fmtScheme.bgFillStyles.map(serializeThemeFill)
    : [defaultSolidFill, defaultSolidFill, defaultSolidFill];

  return {
    tag: 'a:fmtScheme',
    attrs: { name: 'Office' },
    children: [
      { tag: 'a:fillStyleLst', attrs: {} as Record<string, string>, children: fillStyles },
      { tag: 'a:lnStyleLst', attrs: {} as Record<string, string>, children: lnStyles },
      { tag: 'a:effectStyleLst', attrs: {} as Record<string, string>, children: effectStyles },
      { tag: 'a:bgFillStyleLst', attrs: {} as Record<string, string>, children: bgFillStyles },
    ],
  };
}

function serializeThemeFill(fill: FillStyle): ParsedNode {
  if (fill.type === 'solid' && fill.color) {
    return {
      tag: 'a:solidFill',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: fill.color }, children: [] }],
    };
  }
  if (fill.type === 'gradient' && fill.gradientFill) {
    const stops: ParsedNode[] = fill.gradientFill.stops.map(s => ({
      tag: 'a:gs',
      attrs: { pos: String(Math.round(s.position * 1000)) },
      children: [{ tag: 'a:srgbClr', attrs: { val: s.color }, children: [] }],
    }));
    const children: ParsedNode[] = [{ tag: 'a:gsLst', attrs: {} as Record<string, string>, children: stops }];
    if (fill.gradientFill.angle !== undefined) {
      children.push({ tag: 'a:lin', attrs: { ang: String(Math.round(fill.gradientFill.angle * 60000)), scaled: '1' }, children: [] });
    }
    return { tag: 'a:gradFill', attrs: {}, children };
  }
  if (fill.type === 'pattern' && fill.patternFill) {
    const children: ParsedNode[] = [];
    if (fill.patternFill.fgColor) {
      children.push({ tag: 'a:fgClr', attrs: {} as Record<string, string>, children: [{ tag: 'a:srgbClr', attrs: { val: fill.patternFill.fgColor }, children: [] }] });
    }
    if (fill.patternFill.bgColor) {
      children.push({ tag: 'a:bgClr', attrs: {} as Record<string, string>, children: [{ tag: 'a:srgbClr', attrs: { val: fill.patternFill.bgColor }, children: [] }] });
    }
    return { tag: 'a:pattFill', attrs: fill.patternFill.preset ? { prst: fill.patternFill.preset } : {}, children };
  }
  if (fill.type === 'none') {
    return { tag: 'a:noFill', attrs: {} as Record<string, string>, children: [] };
  }
  // Fallback
  const phClr: ParsedNode = { tag: 'a:schemeClr', attrs: { val: 'phClr' }, children: [] };
  return { tag: 'a:solidFill', attrs: {} as Record<string, string>, children: [phClr] };
}

function serializeThemeLine(border: BorderStyle): ParsedNode {
  const lnAttrs: Record<string, string> = {};
  if (border.width) lnAttrs.w = String(Math.round(border.width * 12700));
  if (border.compound) lnAttrs.cmpd = border.compound;
  if (border.cap) lnAttrs.cap = border.cap;

  const lnChildren: ParsedNode[] = [];
  if (border.color) {
    lnChildren.push({
      tag: 'a:solidFill',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: border.color }, children: [] }],
    });
  }
  return { tag: 'a:ln', attrs: lnAttrs, children: lnChildren };
}

function serializeThemeEffect(effect: EffectStyle): ParsedNode {
  const effectChildren: ParsedNode[] = [];

  if (effect.shadow) {
    const shadowAttrs: Record<string, string> = {};
    if (effect.shadow.blur) shadowAttrs.blurRad = String(Math.round(effect.shadow.blur * 12700));
    const shadowChildren: ParsedNode[] = [];
    if (effect.shadow.color) {
      shadowChildren.push({ tag: 'a:srgbClr', attrs: { val: effect.shadow.color }, children: [] });
    }
    effectChildren.push({ tag: 'a:outerShdw', attrs: shadowAttrs, children: shadowChildren });
  }
  if (effect.glow) {
    const glowAttrs: Record<string, string> = {};
    if (effect.glow.radius) glowAttrs.rad = String(Math.round(effect.glow.radius * 12700));
    const glowChildren: ParsedNode[] = [];
    if (effect.glow.color) glowChildren.push({ tag: 'a:srgbClr', attrs: { val: effect.glow.color }, children: [] });
    effectChildren.push({ tag: 'a:glow', attrs: glowAttrs, children: glowChildren });
  }
  if (effect.softEdge) {
    const seAttrs: Record<string, string> = {};
    if (effect.softEdge.radius) seAttrs.rad = String(Math.round(effect.softEdge.radius * 12700));
    effectChildren.push({ tag: 'a:softEdge', attrs: seAttrs, children: [] });
  }

  return {
    tag: 'a:effectStyle',
    attrs: {} as Record<string, string>,
    children: [{ tag: 'a:effectLst', attrs: {} as Record<string, string>, children: effectChildren }],
  };
}
