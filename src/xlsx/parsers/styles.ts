import type { RawDocument, ParsedNode } from '../../core/types.js';
import type {
  CellStyleDefinitions,
  CellStyle,
  FontDefinition,
  FillDefinition,
  BorderDefinition,
  BorderStyle,
  NumberFormatDefinition,
  Alignment,
} from '../types.js';

export function extractStyles(raw: RawDocument): CellStyleDefinitions {
  const stylesXml = raw.parts.get('xl/styles.xml');
  if (!stylesXml) {
    return { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] };
  }

  return {
    numberFormats: extractNumberFormats(stylesXml),
    fonts: extractFonts(stylesXml),
    fills: extractFills(stylesXml),
    borders: extractBorders(stylesXml),
    cellStyles: extractCellXfs(stylesXml),
  };
}

function extractNumberFormats(node: ParsedNode): NumberFormatDefinition[] {
  const numFmts = findChild(node, 'numFmts');
  if (!numFmts) return [];

  const formats: NumberFormatDefinition[] = [];
  for (const child of numFmts.children) {
    if (typeof child === 'string' || child.tag !== 'numFmt') continue;
    const id = child.attrs['numFmtId'];
    const formatCode = child.attrs['formatCode'];
    if (id && formatCode) {
      formats.push({ id, formatCode });
    }
  }
  return formats;
}

function extractFonts(node: ParsedNode): FontDefinition[] {
  const fontsNode = findChild(node, 'fonts');
  if (!fontsNode) return [];

  const fonts: FontDefinition[] = [];
  for (const child of fontsNode.children) {
    if (typeof child === 'string' || child.tag !== 'font') continue;
    fonts.push(parseFont(child));
  }
  return fonts;
}

function parseFont(node: ParsedNode): FontDefinition {
  const font: FontDefinition = {};
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'name': font.name = child.attrs['val']; break;
      case 'sz': font.size = parseFloat(child.attrs['val'] || '0'); break;
      case 'b': font.bold = true; break;
      case 'i': font.italic = true; break;
      case 'u': font.underline = true; break;
      case 'strike': font.strike = true; break;
      case 'color': font.color = parseColor(child); break;
      case 'vertAlign': {
        const v = child.attrs['val'];
        if (v === 'superscript' || v === 'subscript') font.vertAlign = v;
        break;
      }
      case 'scheme': font.scheme = child.attrs['val'] as 'major' | 'minor'; break;
    }
  }
  return font;
}

function extractFills(node: ParsedNode): FillDefinition[] {
  const fillsNode = findChild(node, 'fills');
  if (!fillsNode) return [];

  const fills: FillDefinition[] = [];
  for (const child of fillsNode.children) {
    if (typeof child === 'string' || child.tag !== 'fill') continue;
    fills.push(parseFill(child));
  }
  return fills;
}

function parseFill(node: ParsedNode): FillDefinition {
  const fill: FillDefinition = {};
  const patternFill = findChild(node, 'patternFill');
  if (patternFill) {
    fill.patternType = patternFill.attrs['patternType'];
    const fgColor = findChild(patternFill, 'fgColor');
    if (fgColor) fill.fgColor = parseColor(fgColor);
    const bgColor = findChild(patternFill, 'bgColor');
    if (bgColor) fill.bgColor = parseColor(bgColor);
  }
  const gradientFill = findChild(node, 'gradientFill');
  if (gradientFill) {
    fill.gradientFill = parseGradientFill(gradientFill);
  }
  return fill;
}

function extractBorders(node: ParsedNode): BorderDefinition[] {
  const bordersNode = findChild(node, 'borders');
  if (!bordersNode) return [];

  const borders: BorderDefinition[] = [];
  for (const child of bordersNode.children) {
    if (typeof child === 'string' || child.tag !== 'border') continue;
    borders.push(parseBorder(child));
  }
  return borders;
}

function parseGradientFill(node: ParsedNode): import('../types.js').GradientFill {
  const fill: import('../types.js').GradientFill = { stops: [] };
  if (node.attrs['type']) fill.type = node.attrs['type'] as 'linear' | 'path';
  if (node.attrs['degree']) fill.degree = parseFloat(node.attrs['degree']);
  if (node.attrs['left']) fill.left = parseFloat(node.attrs['left']);
  if (node.attrs['right']) fill.right = parseFloat(node.attrs['right']);
  if (node.attrs['top']) fill.top = parseFloat(node.attrs['top']);
  if (node.attrs['bottom']) fill.bottom = parseFloat(node.attrs['bottom']);

  for (const child of node.children) {
    if (typeof child === 'string' || child.tag !== 'stop') continue;
    const position = parseFloat(child.attrs['position'] || '0');
    const colorNode = findChild(child, 'color');
    const color = colorNode ? parseColor(colorNode) : '';
    fill.stops.push({ position, color });
  }
  return fill;
}

function parseBorder(node: ParsedNode): BorderDefinition {
  const border: BorderDefinition = {};
  for (const side of ['top', 'bottom', 'left', 'right', 'diagonal'] as const) {
    const sideNode = findChild(node, side);
    if (sideNode) {
      border[side] = parseBorderStyle(sideNode);
    }
  }
  if (node.attrs['diagonalUp'] === '1' || node.attrs['diagonalUp'] === 'true') {
    border.diagonalUp = true;
  }
  if (node.attrs['diagonalDown'] === '1' || node.attrs['diagonalDown'] === 'true') {
    border.diagonalDown = true;
  }
  return border;
}

function parseBorderStyle(node: ParsedNode): BorderStyle {
  const style: BorderStyle = {};
  style.style = node.attrs['style'];
  const color = findChild(node, 'color');
  if (color) style.color = parseColor(color);
  return style;
}

function extractCellXfs(node: ParsedNode): CellStyle[] {
  const cellXfs = findChild(node, 'cellXfs');
  if (!cellXfs) return [];

  const styles: CellStyle[] = [];
  for (const child of cellXfs.children) {
    if (typeof child === 'string' || child.tag !== 'xf') continue;
    styles.push(parseXf(child));
  }
  return styles;
}

function parseXf(node: ParsedNode): CellStyle {
  const style: CellStyle = { id: node.attrs['xfId'] || '' };
  if (node.attrs['fontId']) style.font = { /* resolved by index */ };
  if (node.attrs['fillId']) style.fill = { /* resolved by index */ };
  if (node.attrs['borderId']) style.border = { /* resolved by index */ };
  if (node.attrs['numFmtId']) style.numberFormat = node.attrs['numFmtId'];

  const alignment = findChild(node, 'alignment');
  if (alignment) {
    style.alignment = parseAlignment(alignment);
  }

  const protection = findChild(node, 'protection');
  if (protection) {
    style.protection = {
      locked: protection.attrs['locked'] !== '0',
      hidden: protection.attrs['hidden'] === '1',
    };
  }

  return style;
}

function parseAlignment(node: ParsedNode): Alignment {
  const alignment: Alignment = {};
  const horizontal = node.attrs['horizontal'];
  if (horizontal) alignment.horizontal = horizontal as Alignment['horizontal'];
  const vertical = node.attrs['vertical'];
  if (vertical) alignment.vertical = vertical as Alignment['vertical'];
  if (node.attrs['wrapText'] === '1' || node.attrs['wrapText'] === 'true') {
    alignment.wrapText = true;
  }
  if (node.attrs['indent']) alignment.indent = parseInt(node.attrs['indent'], 10);
  if (node.attrs['textRotation']) alignment.textRotation = parseInt(node.attrs['textRotation'], 10);
  if (node.attrs['shrinkToFit'] === '1' || node.attrs['shrinkToFit'] === 'true') {
    alignment.shrinkToFit = true;
  }
  if (node.attrs['readingOrder']) {
    alignment.readingOrder = parseInt(node.attrs['readingOrder'], 10) as 0 | 1 | 2;
  }
  if (node.attrs['justifyLastLine'] === '1' || node.attrs['justifyLastLine'] === 'true') {
    alignment.justifyLastLine = true;
  }
  return alignment;
}

function parseColor(node: ParsedNode): string {
  const rgb = node.attrs['rgb'];
  if (rgb) return rgb;
  const indexed = node.attrs['indexed'];
  if (indexed) return `indexed:${indexed}`;
  const theme = node.attrs['theme'];
  if (theme) return `theme:${theme}`;
  return '';
}

function findChild(node: ParsedNode, tag: string): ParsedNode | undefined {
  for (const child of node.children) {
    if (typeof child !== 'string' && child.tag === tag) {
      return child;
    }
  }
  return undefined;
}
