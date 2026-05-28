import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type {
  Slide, TextShape, ImageShape, MediaShape, ShapeStyle, Hyperlink,
  BodyProperties, FillStyle, BorderStyle, ShadowStyle,
  GlowStyle, SoftEdgeStyle, ReflectionStyle, LineEnd,
} from '../types.js';
import type { Paragraph, TextRun, ParagraphProperties, NumberingProperties } from '../../docx/types.js';
import { NS } from './presentation.js';
import { serializeTransition, serializeTiming } from './animation.js';
import { serializeTableShape } from './table.js';

export function serializeSlide(slide: Slide, _index: number): string {
  // shape ID 计数器，根 spTree 占用 id=1，子 shape 从 2 开始
  const nextId = { value: 2 };
  const spChildren: ParsedNode[] = [];

  for (const element of slide.elements) {
    if (element.type === 'text') {
      spChildren.push(serializeTextShape(element, nextId));
    } else if (element.type === 'image') {
      spChildren.push(serializeImageShape(element, nextId));
    } else if (element.type === 'group') {
      spChildren.push(serializeGroupShape(element, nextId));
    } else if (element.type === 'table') {
      spChildren.push(serializeTableShape(element, nextId));
    } else if (element.type === 'media') {
      spChildren.push(serializeMediaShape(element, nextId));
    }
  }

  const xfrm: ParsedNode = {
    tag: 'a:xfrm',
    attrs: {} as Record<string, string>,
    children: [
      { tag: 'a:off', attrs: { x: '0', y: '0' }, children: [] },
      { tag: 'a:ext', attrs: { cx: '0', cy: '0' }, children: [] },
      { tag: 'a:chOff', attrs: { x: '0', y: '0' }, children: [] },
      { tag: 'a:chExt', attrs: { cx: '0', cy: '0' }, children: [] },
    ],
  };

  const cSldChildren: ParsedNode[] = [];

  // Slide background
  if (slide.background) {
    cSldChildren.push(serializeSlideBackground(slide.background));
  }

  cSldChildren.push({
    tag: 'p:spTree',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:nvGrpSpPr',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'p:cNvPr', attrs: { id: '1', name: '' }, children: [] },
          { tag: 'p:cNvGrpSpPr', attrs: {} as Record<string, string>, children: [] },
          { tag: 'p:nvPr', attrs: {} as Record<string, string>, children: [] },
        ],
      },
      { tag: 'p:grpSpPr', attrs: {} as Record<string, string>, children: [xfrm] },
      ...spChildren,
    ],
  });

  // Build slide attrs
  const sldAttrs: Record<string, string> = {
    'xmlns:p': NS.P_NS, 'xmlns:a': NS.A_NS, 'xmlns:r': NS.R_NS,
  };
  if (slide.showMasterSp === false) sldAttrs.showMasterSp = '0';
  if (slide.showMasterPhAnim === false) sldAttrs.showMasterPhAnim = '0';

  const children: ParsedNode[] = [
    {
      tag: 'p:cSld',
      attrs: {} as Record<string, string>,
      children: cSldChildren,
    },
  ];

  // Color map override
  if (slide.clrMap) {
    const clrMapAttrs: Record<string, string> = {};
    for (const [key, val] of Object.entries(slide.clrMap)) {
      clrMapAttrs[key] = val;
    }
    children.push({
      tag: 'p:clrMapOvr',
      attrs: {} as Record<string, string>,
      children: [{
        tag: 'a:overrideClrMapping',
        attrs: clrMapAttrs,
        children: [],
      }],
    });
  } else {
    children.push({
      tag: 'p:clrMapOvr',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:masterClrMapping', attrs: {} as Record<string, string>, children: [] }],
    });
  }

  if (slide.transition) {
    children.push(serializeTransition(slide.transition));
  }

  if (slide.animations && slide.animations.length > 0) {
    children.push(serializeTiming(slide.animations));
  }

  const root: ParsedNode = {
    tag: 'p:sld',
    attrs: sldAttrs,
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeSlideBackground(bg: FillStyle): ParsedNode {
  return {
    tag: 'p:bg',
    attrs: {} as Record<string, string>,
    children: [{
      tag: 'p:bgPr',
      attrs: {} as Record<string, string>,
      children: serializeFillNodes(bg),
    }],
  };
}

function serializeTextShape(shape: TextShape, nextId?: { value: number }): ParsedNode {
  const pChildren: ParsedNode[] = shape.paragraphs.map(para => serializeParagraph(para));

  const nvPrChildren: ParsedNode[] = [];
  if (shape.placeholder) {
    const phAttrs: Record<string, string> = { type: shape.placeholder.type };
    if (shape.placeholder.index !== undefined) {
      phAttrs.idx = String(shape.placeholder.index);
    }
    nvPrChildren.push({ tag: 'p:ph', attrs: phAttrs, children: [] });
  }

  // spPr children
  const spPrChildren: ParsedNode[] = [];

  // Position with flip/rotation
  const xfrmAttrs: Record<string, string> = {};
  if (shape.position.flipH) xfrmAttrs.flipH = '1';
  if (shape.position.flipV) xfrmAttrs.flipV = '1';
  if (shape.position.rotation) xfrmAttrs.rot = String(Math.round(shape.position.rotation * 60000));

  spPrChildren.push({
    tag: 'a:xfrm',
    attrs: xfrmAttrs,
    children: [
      { tag: 'a:off', attrs: { x: String(shape.position.x), y: String(shape.position.y) }, children: [] },
      { tag: 'a:ext', attrs: { cx: String(shape.position.width), cy: String(shape.position.height) }, children: [] },
    ],
  });

  // Preset geometry
  if (shape.presetGeom) {
    spPrChildren.push({
      tag: 'a:prstGeom',
      attrs: { prst: shape.presetGeom },
      children: [{ tag: 'a:avLst', attrs: {} as Record<string, string>, children: [] }],
    });
  }

  spPrChildren.push(...serializeShapeStyle(shape.style));
  spPrChildren.push(...serializeHyperlink(shape.hyperlink));

  // txBody children
  const txBodyChildren: ParsedNode[] = [];

  // Body properties
  txBodyChildren.push(serializeBodyProperties(shape.bodyProperties));

  // List style
  if (shape.listStyle) {
    txBodyChildren.push(serializeListStyle(shape.listStyle));
  }

  txBodyChildren.push(...pChildren);

  // txBody 必须包含至少一个 a:p
  if (pChildren.length === 0) {
    txBodyChildren.push({ tag: 'a:p', attrs: {} as Record<string, string>, children: [] });
  }

  return {
    tag: 'p:sp',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:nvSpPr',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'p:cNvPr', attrs: { id: String(nextId ? nextId.value++ : 1), name: 'TextBox 1' }, children: [] },
          { tag: 'p:cNvSpPr', attrs: { txBox: '1' }, children: [] },
          { tag: 'p:nvPr', attrs: {} as Record<string, string>, children: nvPrChildren },
        ],
      },
      {
        tag: 'p:spPr',
        attrs: {} as Record<string, string>,
        children: spPrChildren,
      },
      {
        tag: 'p:txBody',
        attrs: {} as Record<string, string>,
        children: txBodyChildren,
      },
    ],
  };
}

function serializeBodyProperties(bp?: BodyProperties): ParsedNode {
  const attrs: Record<string, string> = {};
  if (bp?.anchor) attrs.anchor = bp.anchor;
  if (bp?.wrap) attrs.wrap = bp.wrap;
  if (bp?.leftInset !== undefined) attrs.lIns = String(Math.round(bp.leftInset * 914400));
  if (bp?.topInset !== undefined) attrs.tIns = String(Math.round(bp.topInset * 914400));
  if (bp?.rightInset !== undefined) attrs.rIns = String(Math.round(bp.rightInset * 914400));
  if (bp?.bottomInset !== undefined) attrs.bIns = String(Math.round(bp.bottomInset * 914400));
  if (bp?.vertical) attrs.vert = bp.vertical;

  const children: ParsedNode[] = [];
  if (bp?.autoFit === 'normal') {
    children.push({ tag: 'a:normAutofit', attrs: {} as Record<string, string>, children: [] });
  } else if (bp?.autoFit === 'shape') {
    children.push({ tag: 'a:spAutoFit', attrs: {} as Record<string, string>, children: [] });
  }

  return { tag: 'a:bodyPr', attrs, children };
}

function serializeListStyle(listStyle: NonNullable<TextShape['listStyle']>): ParsedNode {
  const children: ParsedNode[] = [];
  if (listStyle.defaultParagraphProperties) {
    for (const lp of listStyle.defaultParagraphProperties) {
      const tag = `a:lvl${lp.level}pPr`;
      const attrs: Record<string, string> = {};
      if (lp.alignment) attrs.algn = lp.alignment;
      if (lp.marL !== undefined) attrs.marL = String(Math.round(lp.marL * 12700));
      if (lp.indent !== undefined) attrs.indent = String(Math.round(lp.indent * 12700));

      const lvlChildren: ParsedNode[] = [];
      if (lp.fontScale !== undefined) {
        lvlChildren.push({
          tag: 'a:defRPr',
          attrs: { sz: String(Math.round(lp.fontScale * 100)) },
          children: [],
        });
      }

      children.push({ tag, attrs, children: lvlChildren });
    }
  }
  return { tag: 'a:lstStyle', attrs: {} as Record<string, string>, children };
}

function serializeParagraph(para: Paragraph): ParsedNode {
  const children: ParsedNode[] = [];

  // Paragraph properties
  if (para.properties || para.numbering) {
    children.push(serializeParagraphProperties(para.properties, para.numbering));
  }

  // Runs
  for (const run of para.runs) {
    if (run.text === '\n') {
      children.push({ tag: 'a:br', attrs: {} as Record<string, string>, children: [] });
    } else {
      children.push(serializeRun(run));
    }
  }

  return {
    tag: 'a:p',
    attrs: {} as Record<string, string>,
    children,
  };
}

function serializeParagraphProperties(props?: ParagraphProperties, numbering?: NumberingProperties): ParsedNode {
  const attrs: Record<string, string> = {};
  const children: ParsedNode[] = [];

  if (props?.alignment) {
    const map: Record<string, string> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };
    attrs.algn = map[props.alignment] || props.alignment;
  }

  if (props?.indent) {
    if (props.indent.left !== undefined) attrs.marL = String(Math.round(props.indent.left * 12700));
    if (props.indent.right !== undefined) attrs.marR = String(Math.round(props.indent.right * 12700));
    if (props.indent.firstLine !== undefined) attrs.indent = String(Math.round(props.indent.firstLine * 12700));
  }

  // Spacing
  if (props?.spacing) {
    if (props.spacing.before !== undefined) {
      children.push({
        tag: 'a:spcBef',
        attrs: {} as Record<string, string>,
        children: [{ tag: 'a:spcPts', attrs: { val: String(Math.round(props.spacing.before * 100)) }, children: [] }],
      });
    }
    if (props.spacing.after !== undefined) {
      children.push({
        tag: 'a:spcAft',
        attrs: {} as Record<string, string>,
        children: [{ tag: 'a:spcPts', attrs: { val: String(Math.round(props.spacing.after * 100)) }, children: [] }],
      });
    }
    if (props.spacing.line !== undefined) {
      if (props.spacing.lineRule === 'auto') {
        children.push({
          tag: 'a:lnSpc',
          attrs: {} as Record<string, string>,
          children: [{ tag: 'a:spcPct', attrs: { val: String(Math.round(props.spacing.line * 1000)) }, children: [] }],
        });
      } else {
        children.push({
          tag: 'a:lnSpc',
          attrs: {} as Record<string, string>,
          children: [{ tag: 'a:spcPts', attrs: { val: String(Math.round(props.spacing.line * 100)) }, children: [] }],
        });
      }
    }
  }

  // Bullet
  if (numbering) {
    if (numbering.numId === 'buChar' && numbering.format) {
      children.push({ tag: 'a:buChar', attrs: { char: numbering.format }, children: [] });
    } else if (numbering.numId === 'buAutoNum') {
      const buAttrs: Record<string, string> = {};
      if (numbering.format) buAttrs.type = numbering.format;
      if (numbering.text) buAttrs.startAt = numbering.text;
      children.push({ tag: 'a:buAutoNum', attrs: buAttrs, children: [] });
    }
    if (numbering.level !== undefined) attrs.lvl = String(numbering.level);
  }

  return { tag: 'a:pPr', attrs, children };
}

function serializeRun(run: TextRun): ParsedNode {
  const rPrAttrs: Record<string, string> = { lang: 'en-US' };

  if (run.bold) rPrAttrs.b = '1';
  if (run.italic) rPrAttrs.i = '1';
  if (run.underline) rPrAttrs.u = 'sng';
  if (run.strike) rPrAttrs.strike = 'sngStrike';
  if (run.fontSize) rPrAttrs.sz = String(Math.round(run.fontSize * 100));
  if (run.caps) rPrAttrs.cap = 'all';
  if (run.characterSpacing) rPrAttrs.spc = String(Math.round(run.characterSpacing * 100));
  if (run.superscript) rPrAttrs.baseline = '30000';
  if (run.subscript) rPrAttrs.baseline = '-25000';

  const rPrChildren: ParsedNode[] = [];
  if (run.color) {
    rPrChildren.push({
      tag: 'a:solidFill',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: run.color }, children: [] }],
    });
  }
  if (run.fontFamily) {
    rPrChildren.push({ tag: 'a:latin', attrs: { typeface: run.fontFamily }, children: [] });
  }

  return {
    tag: 'a:r',
    attrs: {} as Record<string, string>,
    children: [
      { tag: 'a:rPr', attrs: rPrAttrs, children: rPrChildren },
      { tag: 'a:t', attrs: {} as Record<string, string>, children: [run.text] },
    ],
  };
}

function serializeImageShape(shape: ImageShape, nextId?: { value: number }): ParsedNode {
  const xfrmAttrs: Record<string, string> = {};
  if (shape.position.flipH) xfrmAttrs.flipH = '1';
  if (shape.position.flipV) xfrmAttrs.flipV = '1';
  if (shape.position.rotation) xfrmAttrs.rot = String(Math.round(shape.position.rotation * 60000));

  return {
    tag: 'p:pic',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:nvPicPr',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'p:cNvPr', attrs: { id: String(nextId ? nextId.value++ : 2), name: 'Image 1' }, children: [] },
          { tag: 'p:cNvPicPr', attrs: {} as Record<string, string>, children: [] },
          { tag: 'p:nvPr', attrs: {} as Record<string, string>, children: [] },
        ],
      },
      {
        tag: 'p:blipFill',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'a:blip', attrs: { 'r:embed': shape.relationshipId }, children: [] },
        ],
      },
      {
        tag: 'p:spPr',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'a:xfrm',
            attrs: xfrmAttrs,
            children: [
              { tag: 'a:off', attrs: { x: String(shape.position.x), y: String(shape.position.y) }, children: [] },
              { tag: 'a:ext', attrs: { cx: String(shape.position.width), cy: String(shape.position.height) }, children: [] },
            ],
          },
        ],
      },
    ],
  };
}

function serializeGroupShape(shape: import('../types.js').GroupShape, nextId?: { value: number }): ParsedNode {
  const children: ParsedNode[] = [];
  for (const child of shape.children) {
    if (child.type === 'text') children.push(serializeTextShape(child, nextId));
    else if (child.type === 'image') children.push(serializeImageShape(child, nextId));
    else if (child.type === 'group') children.push(serializeGroupShape(child, nextId));
    else if (child.type === 'table') children.push(serializeTableShape(child, nextId));
    else if (child.type === 'media') children.push(serializeMediaShape(child, nextId));
  }

  const xfrmChildren: ParsedNode[] = [
    { tag: 'a:off', attrs: { x: String(shape.position.x), y: String(shape.position.y) }, children: [] },
    { tag: 'a:ext', attrs: { cx: String(shape.position.width), cy: String(shape.position.height) }, children: [] },
  ];

  if (shape.childOffset) {
    xfrmChildren.push({ tag: 'a:chOff', attrs: { x: String(shape.childOffset.x), y: String(shape.childOffset.y) }, children: [] });
  }
  if (shape.childExtent) {
    xfrmChildren.push({ tag: 'a:chExt', attrs: { cx: String(shape.childExtent.width), cy: String(shape.childExtent.height) }, children: [] });
  }

  return {
    tag: 'p:grpSp',
    attrs: {} as Record<string, string>,
    children: [
      { tag: 'p:nvGrpSpPr', attrs: {} as Record<string, string>, children: [
        { tag: 'p:cNvPr', attrs: { id: String(nextId ? nextId.value++ : 1), name: 'Group' }, children: [] },
        { tag: 'p:cNvGrpSpPr', attrs: {} as Record<string, string>, children: [] },
        { tag: 'p:nvPr', attrs: {} as Record<string, string>, children: [] },
      ]},
      { tag: 'p:grpSpPr', attrs: {} as Record<string, string>, children: [
        { tag: 'a:xfrm', attrs: {} as Record<string, string>, children: xfrmChildren },
      ]},
      ...children,
    ],
  };
}

function serializeShapeStyle(style?: ShapeStyle): ParsedNode[] {
  if (!style) return [];
  const children: ParsedNode[] = [];

  if (style.fill) {
    children.push(...serializeFillNodes(style.fill));
  }

  if (style.border) {
    children.push(serializeBorder(style.border));
  }

  if (style.shadow || style.glow || style.softEdge || style.reflection) {
    children.push(serializeEffects(style));
  }

  return children;
}

function serializeFillNodes(fill: FillStyle): ParsedNode[] {
  if (fill.type === 'solid' && fill.color) {
    return [{
      tag: 'a:solidFill',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: fill.color }, children: [] }],
    }];
  }
  if (fill.type === 'none') {
    return [{ tag: 'a:noFill', attrs: {} as Record<string, string>, children: [] }];
  }
  if (fill.type === 'gradient' && fill.gradientFill) {
    return [serializeGradientFill(fill.gradientFill)];
  }
  if (fill.type === 'pattern' && fill.patternFill) {
    return [serializePatternFill(fill.patternFill)];
  }
  if (fill.type === 'blip' && fill.blipRelationshipId) {
    return [{
      tag: 'a:blipFill',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:blip', attrs: { 'r:embed': fill.blipRelationshipId }, children: [] }],
    }];
  }
  if (fill.type === 'group') {
    return [{ tag: 'a:grpFill', attrs: {} as Record<string, string>, children: [] }];
  }
  return [];
}

function serializeGradientFill(gf: import('../types.js').GradientFill): ParsedNode {
  const children: ParsedNode[] = [];

  if (gf.angle !== undefined) {
    children.push({
      tag: 'a:lin',
      attrs: { ang: String(Math.round(gf.angle * 60000)), scaled: '1' },
      children: [],
    });
  }

  const stops: ParsedNode[] = gf.stops.map(s => ({
    tag: 'a:gs',
    attrs: { pos: String(Math.round(s.position * 1000)) },
    children: [{ tag: 'a:srgbClr', attrs: { val: s.color }, children: [] }],
  }));
  children.push({ tag: 'a:gsLst', attrs: {} as Record<string, string>, children: stops });

  const attrs: Record<string, string> = {};
  if (gf.type) attrs.type = gf.type;

  return { tag: 'a:gradFill', attrs, children };
}

function serializePatternFill(pf: import('../types.js').PatternFill): ParsedNode {
  const children: ParsedNode[] = [];
  if (pf.fgColor) {
    children.push({
      tag: 'a:fgClr',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: pf.fgColor }, children: [] }],
    });
  }
  if (pf.bgColor) {
    children.push({
      tag: 'a:bgClr',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: pf.bgColor }, children: [] }],
    });
  }
  const attrs: Record<string, string> = {};
  if (pf.preset) attrs.prst = pf.preset;
  return { tag: 'a:pattFill', attrs, children };
}

function serializeBorder(border: BorderStyle): ParsedNode {
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
  const dashVal = border.dashType || (border.style === 'dashed' ? 'dash' : border.style === 'dotted' ? 'dot' : undefined);
  if (dashVal) {
    lnChildren.push({ tag: 'a:prstDash', attrs: { val: dashVal }, children: [] });
  }
  if (border.headEnd) lnChildren.push(serializeLineEnd('a:headEnd', border.headEnd));
  if (border.tailEnd) lnChildren.push(serializeLineEnd('a:tailEnd', border.tailEnd));

  return { tag: 'a:ln', attrs: lnAttrs, children: lnChildren };
}

function serializeLineEnd(tag: string, end: LineEnd): ParsedNode {
  const attrs: Record<string, string> = {};
  if (end.type) attrs.type = end.type;
  if (end.width) attrs.w = end.width;
  if (end.length) attrs.len = end.length;
  return { tag, attrs, children: [] };
}

function serializeEffects(style: ShapeStyle): ParsedNode {
  const effectChildren: ParsedNode[] = [];

  if (style.shadow) {
    const shadowAttrs: Record<string, string> = {};
    if (style.shadow.blur) shadowAttrs.blurRad = String(Math.round(style.shadow.blur * 12700));
    if (style.shadow.offsetX || style.shadow.offsetY) {
      const dist = Math.sqrt((style.shadow.offsetX || 0) ** 2 + (style.shadow.offsetY || 0) ** 2);
      shadowAttrs.dist = String(Math.round(dist * 12700));
      const dir = Math.atan2(style.shadow.offsetY || 0, style.shadow.offsetX || 0);
      shadowAttrs.dir = String(Math.round((dir * 180 / Math.PI + 360) % 360 * 60000));
    }
    if (style.shadow.color) {
      shadowAttrs.algn = 'bl';
    }

    const shadowChildren: ParsedNode[] = [];
    if (style.shadow.color) {
      shadowChildren.push({ tag: 'a:srgbClr', attrs: { val: style.shadow.color }, children: [] });
    }

    effectChildren.push({
      tag: style.shadow.type === 'inner' ? 'a:innerShdw' : 'a:outerShdw',
      attrs: shadowAttrs,
      children: shadowChildren,
    });
  }

  if (style.glow) {
    const glowAttrs: Record<string, string> = {};
    if (style.glow.radius) glowAttrs.rad = String(Math.round(style.glow.radius * 12700));
    const glowChildren: ParsedNode[] = [];
    if (style.glow.color) {
      glowChildren.push({ tag: 'a:srgbClr', attrs: { val: style.glow.color }, children: [] });
    }
    effectChildren.push({ tag: 'a:glow', attrs: glowAttrs, children: glowChildren });
  }

  if (style.softEdge) {
    const seAttrs: Record<string, string> = {};
    if (style.softEdge.radius) seAttrs.rad = String(Math.round(style.softEdge.radius * 12700));
    effectChildren.push({ tag: 'a:softEdge', attrs: seAttrs, children: [] });
  }

  if (style.reflection) {
    effectChildren.push(serializeReflection(style.reflection));
  }

  return {
    tag: 'a:effectLst',
    attrs: {} as Record<string, string>,
    children: effectChildren,
  };
}

function serializeReflection(r: ReflectionStyle): ParsedNode {
  const attrs: Record<string, string> = {};
  if (r.blur !== undefined) attrs.blurRad = String(Math.round(r.blur * 12700));
  if (r.distance !== undefined) attrs.dist = String(Math.round(r.distance * 12700));
  if (r.startOpacity !== undefined) attrs.stA = String(Math.round(r.startOpacity * 1000));
  if (r.endOpacity !== undefined) attrs.endA = String(Math.round(r.endOpacity * 1000));
  if (r.direction !== undefined) attrs.dir = String(Math.round(r.direction * 60000));
  if (r.scaleY !== undefined) attrs.sy = String(Math.round(r.scaleY * 1000));
  return { tag: 'a:reflection', attrs, children: [] };
}

function serializeMediaShape(shape: MediaShape, nextId?: { value: number }): ParsedNode {
  const mediaTag = shape.mediaType === 'video' ? 'a:videoFile' : 'a:audioFile';

  return {
    tag: 'p:pic',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:nvPicPr',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'p:cNvPr', attrs: { id: String(nextId ? nextId.value++ : 2), name: `${shape.mediaType} 1` }, children: [] },
          { tag: 'p:cNvPicPr', attrs: {} as Record<string, string>, children: [] },
          {
            tag: 'p:nvPr',
            attrs: {} as Record<string, string>,
            children: [{ tag: mediaTag, attrs: { 'r:link': shape.relationshipId }, children: [] }],
          },
        ],
      },
      { tag: 'p:blipFill', attrs: {} as Record<string, string>, children: [] },
      {
        tag: 'p:spPr',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'a:xfrm',
            attrs: {} as Record<string, string>,
            children: [
              { tag: 'a:off', attrs: { x: String(shape.position.x), y: String(shape.position.y) }, children: [] },
              { tag: 'a:ext', attrs: { cx: String(shape.position.width), cy: String(shape.position.height) }, children: [] },
            ],
          },
        ],
      },
    ],
  };
}

function serializeHyperlink(hyperlink?: Hyperlink): ParsedNode[] {
  if (!hyperlink) return [];

  const attrs: Record<string, string> = {
    'r:id': hyperlink.url,
  };
  if (hyperlink.tooltip) {
    attrs['tooltip'] = hyperlink.tooltip;
  }

  return [{
    tag: 'a:hlinkClick',
    attrs,
    children: [],
  }];
}
