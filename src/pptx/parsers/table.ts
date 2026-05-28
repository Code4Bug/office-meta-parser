import type { ParsedNode } from '../../core/types.js';
import type { TableShape, PptxTableRow, PptxTableCell, PptxTableCellStyle, Position, FillStyle, BorderStyle } from '../types.js';
import type { Paragraph, TextRun } from '../../docx/types.js';
import { findChild, getTextContent } from './utils.js';
import { parseFill, parseLineNode } from './style.js';

export function parseTableShape(node: ParsedNode): TableShape | null {
  const graphic = findChild(node, 'a:graphic');
  if (!graphic) return null;
  const graphicData = findChild(graphic, 'a:graphicData');
  if (!graphicData) return null;
  const tbl = findChild(graphicData, 'a:tbl');
  if (!tbl) return null;

  // Extract position from the shape's spPr
  const spPr = findChild(node, 'p:spPr');
  const position = spPr ? parsePosition(spPr) : { x: 0, y: 0, width: 0, height: 0 };

  // Table style
  let tableStyleId: string | undefined;
  const tblPr = findChild(tbl, 'a:tblPr');
  if (tblPr?.attrs['styleId']) tableStyleId = tblPr.attrs['styleId'];

  const rows: PptxTableRow[] = [];
  for (const child of tbl.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:tr') {
      rows.push(parseTableRow(child));
    }
  }

  const shape: TableShape = { type: 'table', position, rows };
  if (tableStyleId) shape.tableStyleId = tableStyleId;
  return shape;
}

function parseTableRow(node: ParsedNode): PptxTableRow {
  const cells: PptxTableCell[] = [];
  const height = node.attrs['h'] ? parseInt(node.attrs['h'], 10) : undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:tc') {
      cells.push(parseTableCell(child));
    }
  }

  const row: PptxTableRow = { cells };
  if (height) row.height = height;
  return row;
}

function parseTableCell(node: ParsedNode): PptxTableCell {
  const content: Paragraph[] = [];
  const width = node.attrs['w'] ? parseInt(node.attrs['w'], 10) : undefined;
  const colspan = node.attrs['gridSpan'] ? parseInt(node.attrs['gridSpan'], 10) : undefined;
  const rowspan = node.attrs['rowSpan'] ? parseInt(node.attrs['rowSpan'], 10) : undefined;

  const txBody = findChild(node, 'a:txBody');
  if (txBody) {
    for (const child of txBody.children) {
      if (typeof child === 'string') continue;
      if (child.tag === 'a:p') {
        content.push(parseParagraph(child));
      }
    }
  }

  const cell: PptxTableCell = { content };
  if (width) cell.width = width;
  if (colspan && colspan > 1) cell.colspan = colspan;
  if (rowspan && rowspan > 1) cell.rowspan = rowspan;

  // Cell properties (fill, border)
  const tcPr = findChild(node, 'a:tcPr');
  if (tcPr) {
    const cellStyle: PptxTableCellStyle = {};
    const fill = parseFill(tcPr);
    if (fill) cellStyle.fill = fill;
    const borders = parseCellBorders(tcPr);
    if (borders) cellStyle.borders = borders;
    if (Object.keys(cellStyle).length > 0) cell.style = cellStyle;
  }

  return cell;
}

function parseCellBorders(tcPr: ParsedNode): PptxTableCellStyle['borders'] | undefined {
  const lnL = findChild(tcPr, 'a:lnL');
  const lnR = findChild(tcPr, 'a:lnR');
  const lnT = findChild(tcPr, 'a:lnT');
  const lnB = findChild(tcPr, 'a:lnB');
  if (!lnL && !lnR && !lnT && !lnB) return undefined;

  const borders: PptxTableCellStyle['borders'] = {};
  if (lnL) borders.left = parseLineNode(lnL);
  if (lnR) borders.right = parseLineNode(lnR);
  if (lnT) borders.top = parseLineNode(lnT);
  if (lnB) borders.bottom = parseLineNode(lnB);
  return borders;
}

function parseParagraph(node: ParsedNode): Paragraph {
  const runs: TextRun[] = [];

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:r') {
      runs.push(parseRun(child));
    }
  }

  return { type: 'paragraph', runs };
}

function parseRun(node: ParsedNode): TextRun {
  let text = '';

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:t') {
      text = getTextContent(child);
    }
  }

  return { text };
}

function parsePosition(spPr: ParsedNode): Position {
  const xfrm = findChild(spPr, 'a:xfrm');
  if (!xfrm) return { x: 0, y: 0, width: 0, height: 0 };

  const off = findChild(xfrm, 'a:off');
  const ext = findChild(xfrm, 'a:ext');

  return {
    x: off?.attrs['x'] ? parseInt(off.attrs['x'], 10) : 0,
    y: off?.attrs['y'] ? parseInt(off.attrs['y'], 10) : 0,
    width: ext?.attrs['cx'] ? parseInt(ext.attrs['cx'], 10) : 0,
    height: ext?.attrs['cy'] ? parseInt(ext.attrs['cy'], 10) : 0,
  };
}
