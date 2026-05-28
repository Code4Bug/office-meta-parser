import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type {
  CellStyleDefinitions,
  FontDefinition,
  FillDefinition,
  BorderDefinition,
  CellStyle,
  GradientFill,
} from '../types.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export function serializeCellStyles(styles: CellStyleDefinitions): string {
  const children: ParsedNode[] = [];

  // Number formats
  if (styles.numberFormats.length > 0) {
    const numFmts = styles.numberFormats.map(nf => ({
      tag: 'numFmt',
      attrs: { numFmtId: nf.id, formatCode: nf.formatCode },
      children: [],
    }));
    children.push({ tag: 'numFmts', attrs: { count: String(numFmts.length) }, children: numFmts });
  }

  // Fonts
  const fonts = styles.fonts.map(f => serializeFont(f));
  children.push({ tag: 'fonts', attrs: { count: String(fonts.length) }, children: fonts });

  // Fills
  const fills = styles.fills.map(f => serializeFill(f));
  children.push({ tag: 'fills', attrs: { count: String(fills.length) }, children: fills });

  // Borders
  const borders = styles.borders.map(b => serializeBorder(b));
  children.push({ tag: 'borders', attrs: { count: String(borders.length) }, children: borders });

  // Cell style Xfs
  const cellXfs = styles.cellStyles.map(cs => serializeCellStyle(cs));
  children.push({ tag: 'cellXfs', attrs: { count: String(cellXfs.length) }, children: cellXfs });

  const root: ParsedNode = {
    tag: 'styleSheet',
    attrs: { xmlns: MAIN_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeFont(font: FontDefinition): ParsedNode {
  const children: ParsedNode[] = [];

  if (font.name) {
    children.push({ tag: 'name', attrs: { val: font.name }, children: [] });
  }
  if (font.size) {
    children.push({ tag: 'sz', attrs: { val: String(font.size) }, children: [] });
  }
  if (font.bold) {
    children.push({ tag: 'b', attrs: {}, children: [] });
  }
  if (font.italic) {
    children.push({ tag: 'i', attrs: {}, children: [] });
  }
  if (font.underline) {
    children.push({ tag: 'u', attrs: {}, children: [] });
  }
  if (font.strike) {
    children.push({ tag: 'strike', attrs: {}, children: [] });
  }
  if (font.vertAlign) {
    children.push({ tag: 'vertAlign', attrs: { val: font.vertAlign }, children: [] });
  }
  if (font.color) {
    children.push({ tag: 'color', attrs: { rgb: font.color }, children: [] });
  }
  if (font.scheme) {
    children.push({ tag: 'scheme', attrs: { val: font.scheme }, children: [] });
  }

  return { tag: 'font', attrs: {}, children };
}

function serializeFill(fill: FillDefinition): ParsedNode {
  const children: ParsedNode[] = [];

  if (fill.patternType) {
    const patternFill: ParsedNode = {
      tag: 'patternFill',
      attrs: { patternType: fill.patternType },
      children: [],
    };

    if (fill.fgColor) {
      patternFill.children.push({ tag: 'fgColor', attrs: { rgb: fill.fgColor }, children: [] });
    }
    if (fill.bgColor) {
      patternFill.children.push({ tag: 'bgColor', attrs: { rgb: fill.bgColor }, children: [] });
    }

    children.push(patternFill);
  }

  if (fill.gradientFill) {
    children.push(serializeGradientFill(fill.gradientFill));
  }

  return { tag: 'fill', attrs: {}, children };
}

function serializeGradientFill(gf: GradientFill): ParsedNode {
  const attrs: Record<string, string> = {};
  if (gf.type) attrs.type = gf.type;
  if (gf.degree !== undefined) attrs.degree = String(gf.degree);
  if (gf.left !== undefined) attrs.left = String(gf.left);
  if (gf.right !== undefined) attrs.right = String(gf.right);
  if (gf.top !== undefined) attrs.top = String(gf.top);
  if (gf.bottom !== undefined) attrs.bottom = String(gf.bottom);

  const stops = gf.stops.map(stop => ({
    tag: 'stop',
    attrs: { position: String(stop.position) },
    children: [{ tag: 'color', attrs: { rgb: stop.color }, children: [] }],
  }));

  return { tag: 'gradientFill', attrs, children: stops };
}

function serializeBorder(border: BorderDefinition): ParsedNode {
  const children: ParsedNode[] = [];
  const attrs: Record<string, string> = {};

  if (border.top) {
    children.push(serializeBorderStyle('top', border.top));
  }
  if (border.bottom) {
    children.push(serializeBorderStyle('bottom', border.bottom));
  }
  if (border.left) {
    children.push(serializeBorderStyle('left', border.left));
  }
  if (border.right) {
    children.push(serializeBorderStyle('right', border.right));
  }
  if (border.diagonal) {
    children.push(serializeBorderStyle('diagonal', border.diagonal));
  }
  if (border.diagonalUp) attrs.diagonalUp = '1';
  if (border.diagonalDown) attrs.diagonalDown = '1';

  return { tag: 'border', attrs, children };
}

function serializeBorderStyle(tag: string, style: { style?: string; color?: string }): ParsedNode {
  const attrs: Record<string, string> = {};
  if (style.style) {
    attrs.style = style.style;
  }

  const children: ParsedNode[] = [];
  if (style.color) {
    children.push({ tag: 'color', attrs: { rgb: style.color }, children: [] });
  }

  return { tag, attrs, children };
}

function serializeCellStyle(style: CellStyle): ParsedNode {
  const attrs: Record<string, string> = {};

  if (style.numberFormat) {
    attrs.numFmtId = style.numberFormat;
  }

  const children: ParsedNode[] = [];

  if (style.alignment) {
    const alignmentAttrs: Record<string, string> = {};
    if (style.alignment.horizontal) {
      alignmentAttrs.horizontal = style.alignment.horizontal;
    }
    if (style.alignment.vertical) {
      alignmentAttrs.vertical = style.alignment.vertical;
    }
    if (style.alignment.wrapText) {
      alignmentAttrs.wrapText = '1';
    }
    if (style.alignment.indent !== undefined) {
      alignmentAttrs.indent = String(style.alignment.indent);
    }
    if (style.alignment.textRotation !== undefined) {
      alignmentAttrs.textRotation = String(style.alignment.textRotation);
    }
    if (style.alignment.shrinkToFit) {
      alignmentAttrs.shrinkToFit = '1';
    }
    if (style.alignment.readingOrder !== undefined) {
      alignmentAttrs.readingOrder = String(style.alignment.readingOrder);
    }
    if (style.alignment.justifyLastLine) {
      alignmentAttrs.justifyLastLine = '1';
    }
    children.push({ tag: 'alignment', attrs: alignmentAttrs, children: [] });
  }

  if (style.protection) {
    const protAttrs: Record<string, string> = {
      locked: style.protection.locked === false ? '0' : '1',
      hidden: style.protection.hidden ? '1' : '0',
    };
    children.push({ tag: 'protection', attrs: protAttrs, children: [] });
  }

  return { tag: 'xf', attrs, children };
}
