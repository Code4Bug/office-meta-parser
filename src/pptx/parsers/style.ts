import type { ParsedNode } from '../../core/types.js';
import type {
  ShapeStyle, FillStyle, BorderStyle, ShadowStyle,
  GradientFill, GradientStop, PatternFill, LineEnd,
  GlowStyle, SoftEdgeStyle, ReflectionStyle,
} from '../types.js';
import { findChild } from './utils.js';

export function parseShapeStyle(spPr: ParsedNode): ShapeStyle | undefined {
  const fill = parseFill(spPr);
  const border = parseBorder(spPr);
  const shadow = parseShadow(spPr);
  const opacity = parseOpacity(spPr);
  const glow = parseGlow(spPr);
  const softEdge = parseSoftEdge(spPr);
  const reflection = parseReflection(spPr);

  if (!fill && !border && !shadow && opacity === undefined && !glow && !softEdge && !reflection) return undefined;

  const style: ShapeStyle = {};
  if (fill) style.fill = fill;
  if (border) style.border = border;
  if (shadow) style.shadow = shadow;
  if (opacity !== undefined) style.opacity = opacity;
  if (glow) style.glow = glow;
  if (softEdge) style.softEdge = softEdge;
  if (reflection) style.reflection = reflection;
  return style;
}

export function parseFill(spPr: ParsedNode): FillStyle | undefined {
  const solidFill = findChild(spPr, 'a:solidFill');
  if (solidFill) {
    const color = parseColorValue(solidFill);
    if (color) return { type: 'solid', color };
  }

  const gradFill = findChild(spPr, 'a:gradFill');
  if (gradFill) return { type: 'gradient', gradientFill: parseGradientFill(gradFill) };

  const pattFill = findChild(spPr, 'a:pattFill');
  if (pattFill) return { type: 'pattern', patternFill: parsePatternFill(pattFill) };

  const blipFill = findChild(spPr, 'a:blipFill');
  if (blipFill) {
    const blip = findChild(blipFill, 'a:blip');
    if (blip && blip.attrs['r:embed']) {
      return { type: 'blip', blipRelationshipId: blip.attrs['r:embed'] };
    }
  }

  const noFill = findChild(spPr, 'a:noFill');
  if (noFill) return { type: 'none' };

  const grpFill = findChild(spPr, 'a:grpFill');
  if (grpFill) return { type: 'group' };

  return undefined;
}

export function parseGradientFill(node: ParsedNode): GradientFill {
  const gf: GradientFill = { stops: [] };
  if (node.attrs['type']) gf.type = node.attrs['type'] as 'linear' | 'path';
  if (node.attrs['rotWithShape'] === '0') { /* no rotation */ }

  // Parse linear fill angle
  const lin = findChild(node, 'a:lin');
  if (lin && lin.attrs['ang']) {
    gf.angle = parseInt(lin.attrs['ang'], 10) / 60000;
  }

  // Parse gradient stops
  const gsLst = findChild(node, 'a:gsLst');
  if (gsLst) {
    for (const child of gsLst.children) {
      if (typeof child === 'string' || child.tag !== 'a:gs') continue;
      const pos = child.attrs['pos'] ? parseInt(child.attrs['pos'], 10) / 1000 : 0;
      const color = parseColorValue(child);
      if (color) gf.stops.push({ position: pos, color });
    }
  }

  return gf;
}

function parsePatternFill(node: ParsedNode): PatternFill {
  const pf: PatternFill = {};
  if (node.attrs['prst']) pf.preset = node.attrs['prst'];
  const fgClr = findChild(node, 'a:fgClr');
  if (fgClr) pf.fgColor = parseColorValue(fgClr);
  const bgClr = findChild(node, 'a:bgClr');
  if (bgClr) pf.bgColor = parseColorValue(bgClr);
  return pf;
}

export function parseLineNode(ln: ParsedNode): BorderStyle {
  const border: BorderStyle = {};

  if (ln.attrs['w']) {
    border.width = parseInt(ln.attrs['w'], 10) / 12700;
  }

  const solidFill = findChild(ln, 'a:solidFill');
  if (solidFill) {
    border.color = parseColorValue(solidFill);
  }

  const prstDash = findChild(ln, 'a:prstDash');
  if (prstDash) {
    const val = prstDash.attrs['val'];
    border.dashType = val;
    if (val === 'dash' || val === 'lgDash' || val === 'dashDot' || val === 'lgDashDot' || val === 'lgDashDotDot') {
      border.style = 'dashed';
    } else if (val === 'dot' || val === 'sysDot' || val === 'sysDash') {
      border.style = 'dotted';
    } else {
      border.style = 'solid';
    }
  }

  const headEnd = findChild(ln, 'a:headEnd');
  if (headEnd) border.headEnd = parseLineEnd(headEnd);
  const tailEnd = findChild(ln, 'a:tailEnd');
  if (tailEnd) border.tailEnd = parseLineEnd(tailEnd);

  if (ln.attrs['cmpd']) border.compound = ln.attrs['cmpd'] as BorderStyle['compound'];
  if (ln.attrs['cap']) border.cap = ln.attrs['cap'] as BorderStyle['cap'];

  return border;
}

function parseBorder(spPr: ParsedNode): BorderStyle | undefined {
  const ln = findChild(spPr, 'a:ln');
  if (!ln) return undefined;
  const border = parseLineNode(ln);
  return Object.keys(border).length > 0 ? border : undefined;
}

function parseLineEnd(node: ParsedNode): LineEnd {
  const end: LineEnd = {};
  if (node.attrs['type']) end.type = node.attrs['type'] as LineEnd['type'];
  if (node.attrs['w']) end.width = node.attrs['w'] as LineEnd['width'];
  if (node.attrs['len']) end.length = node.attrs['len'] as LineEnd['length'];
  return end;
}

function parseShadow(spPr: ParsedNode): ShadowStyle | undefined {
  const effectLst = findChild(spPr, 'a:effectLst');
  if (!effectLst) return undefined;

  const outerShdw = findChild(effectLst, 'a:outerShdw');
  if (outerShdw) return parseShadowProps(outerShdw, 'outer');

  const innerShdw = findChild(effectLst, 'a:innerShdw');
  if (innerShdw) return parseShadowProps(innerShdw, 'inner');

  return undefined;
}

function parseShadowProps(node: ParsedNode, type: 'outer' | 'inner'): ShadowStyle {
  const shadow: ShadowStyle = { type };

  if (node.attrs['blurRad']) {
    shadow.blur = parseInt(node.attrs['blurRad'], 10) / 12700;
  }
  if (node.attrs['dist']) {
    const dist = parseInt(node.attrs['dist'], 10) / 12700;
    const dir = node.attrs['dir'] ? parseInt(node.attrs['dir'], 10) / 60000 : 0;
    shadow.offsetX = Math.round(dist * Math.cos(dir * Math.PI / 180));
    shadow.offsetY = Math.round(dist * Math.sin(dir * Math.PI / 180));
  }

  const color = parseColorValue(node);
  if (color) shadow.color = color;

  return shadow;
}

function parseOpacity(spPr: ParsedNode): number | undefined {
  const solidFill = findChild(spPr, 'a:solidFill');
  if (!solidFill) return undefined;

  const srgbClr = findChild(solidFill, 'a:srgbClr');
  if (srgbClr) {
    const alpha = findChild(srgbClr, 'a:alpha');
    if (alpha && alpha.attrs['val']) {
      return parseInt(alpha.attrs['val'], 10) / 100000;
    }
  }

  return undefined;
}

function parseGlow(spPr: ParsedNode): GlowStyle | undefined {
  const effectLst = findChild(spPr, 'a:effectLst');
  if (!effectLst) return undefined;

  const glow = findChild(effectLst, 'a:glow');
  if (!glow) return undefined;

  const gs: GlowStyle = {};
  if (glow.attrs['rad']) gs.radius = parseInt(glow.attrs['rad'], 10) / 12700;
  const color = parseColorValue(glow);
  if (color) gs.color = color;
  return gs;
}

function parseSoftEdge(spPr: ParsedNode): SoftEdgeStyle | undefined {
  const effectLst = findChild(spPr, 'a:effectLst');
  if (!effectLst) return undefined;

  const softEdge = findChild(effectLst, 'a:softEdge');
  if (!softEdge) return undefined;

  const se: SoftEdgeStyle = {};
  if (softEdge.attrs['rad']) se.radius = parseInt(softEdge.attrs['rad'], 10) / 12700;
  return se;
}

function parseReflection(spPr: ParsedNode): ReflectionStyle | undefined {
  const effectLst = findChild(spPr, 'a:effectLst');
  if (!effectLst) return undefined;

  const refl = findChild(effectLst, 'a:reflection');
  if (!refl) return undefined;

  const r: ReflectionStyle = {};
  if (refl.attrs['blurRad']) r.blur = parseInt(refl.attrs['blurRad'], 10) / 12700;
  if (refl.attrs['dist']) r.distance = parseInt(refl.attrs['dist'], 10) / 12700;
  if (refl.attrs['stA']) r.startOpacity = parseInt(refl.attrs['stA'], 10) / 1000;
  if (refl.attrs['endA']) r.endOpacity = parseInt(refl.attrs['endA'], 10) / 1000;
  if (refl.attrs['dir']) r.direction = parseInt(refl.attrs['dir'], 10) / 60000;
  if (refl.attrs['sy']) r.scaleY = parseInt(refl.attrs['sy'], 10) / 1000;
  return Object.keys(r).length > 0 ? r : undefined;
}

export function parseColorValue(parent: ParsedNode): string | undefined {
  const srgbClr = findChild(parent, 'a:srgbClr');
  if (srgbClr && srgbClr.attrs['val']) return srgbClr.attrs['val'];

  const schemeClr = findChild(parent, 'a:schemeClr');
  if (schemeClr && schemeClr.attrs['val']) return `scheme:${schemeClr.attrs['val']}`;

  const prstClr = findChild(parent, 'a:prstClr');
  if (prstClr && prstClr.attrs['val']) return prstClr.attrs['val'];

  return undefined;
}
