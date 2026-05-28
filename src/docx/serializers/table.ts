import type { ParsedNode } from '../../core/types.js';
import type { Table, TableRow, TableCell } from '../types.js';
import { serializeParagraph } from './paragraph.js';
import { serializeImageParagraph } from './image.js';

export function serializeTable(table: Table): ParsedNode {
  const children: ParsedNode[] = [];

  if (table.properties) {
    const tblPrChildren: ParsedNode[] = [];
    if (table.properties.style) {
      tblPrChildren.push({ tag: 'w:tblStyle', attrs: { 'w:val': table.properties.style }, children: [] });
    }
    if (table.properties.width) {
      const widthType = table.properties.widthType || 'dxa';
      tblPrChildren.push({ tag: 'w:tblW', attrs: { 'w:w': String(table.properties.width), 'w:type': widthType }, children: [] });
    }
    if (table.properties.indent !== undefined) {
      tblPrChildren.push({ tag: 'w:tblInd', attrs: { 'w:w': String(table.properties.indent), 'w:type': 'dxa' }, children: [] });
    }
    if (table.properties.borders) {
      const borderChildren: ParsedNode[] = [];
      const b = table.properties.borders;
      if (b.top) borderChildren.push(serializeBorderStyle('w:top', b.top));
      if (b.left) borderChildren.push(serializeBorderStyle('w:left', b.left));
      if (b.bottom) borderChildren.push(serializeBorderStyle('w:bottom', b.bottom));
      if (b.right) borderChildren.push(serializeBorderStyle('w:right', b.right));
      if (b.insideHorizontal) borderChildren.push(serializeBorderStyle('w:insideH', b.insideHorizontal));
      if (b.insideVertical) borderChildren.push(serializeBorderStyle('w:insideV', b.insideVertical));
      tblPrChildren.push({ tag: 'w:tblBorders', attrs: {}, children: borderChildren });
    }
    if (table.properties.layout) {
      tblPrChildren.push({ tag: 'w:tblLayout', attrs: { 'w:type': table.properties.layout }, children: [] });
    }
    if (table.properties.cellMarginTop !== undefined || table.properties.cellMarginLeft !== undefined ||
        table.properties.cellMarginBottom !== undefined || table.properties.cellMarginRight !== undefined) {
      const cellMarChildren: ParsedNode[] = [];
      if (table.properties.cellMarginTop !== undefined) {
        cellMarChildren.push({ tag: 'w:top', attrs: { 'w:w': String(table.properties.cellMarginTop), 'w:type': 'dxa' }, children: [] });
      }
      if (table.properties.cellMarginLeft !== undefined) {
        cellMarChildren.push({ tag: 'w:left', attrs: { 'w:w': String(table.properties.cellMarginLeft), 'w:type': 'dxa' }, children: [] });
      }
      if (table.properties.cellMarginBottom !== undefined) {
        cellMarChildren.push({ tag: 'w:bottom', attrs: { 'w:w': String(table.properties.cellMarginBottom), 'w:type': 'dxa' }, children: [] });
      }
      if (table.properties.cellMarginRight !== undefined) {
        cellMarChildren.push({ tag: 'w:right', attrs: { 'w:w': String(table.properties.cellMarginRight), 'w:type': 'dxa' }, children: [] });
      }
      tblPrChildren.push({ tag: 'w:tblCellMar', attrs: {}, children: cellMarChildren });
    }
    if (tblPrChildren.length > 0) {
      children.push({ tag: 'w:tblPr', attrs: {}, children: tblPrChildren });
    }
  }

  // 添加表格网格列定义
  if (table.properties?.gridColumns && table.properties.gridColumns.length > 0) {
    const gridChildren = table.properties.gridColumns.map(w =>
      ({ tag: 'w:gridCol', attrs: { 'w:w': String(w) }, children: [] })
    );
    children.push({ tag: 'w:tblGrid', attrs: {}, children: gridChildren });
  }

  for (const row of table.rows) {
    children.push(serializeTableRow(row));
  }

  return { tag: 'w:tbl', attrs: {}, children };
}

function serializeBorderStyle(tag: string, border: { style: string; size?: number; color?: string }): ParsedNode {
  const attrs: Record<string, string> = { 'w:val': border.style };
  if (border.size) attrs['w:sz'] = String(border.size);
  if (border.color) attrs['w:color'] = border.color;
  attrs['w:space'] = '0';
  return { tag, attrs, children: [] };
}

function serializeTableRow(row: TableRow): ParsedNode {
  const children: ParsedNode[] = [];

  if (row.properties?.height) {
    const trHeightAttrs: Record<string, string> = { 'w:val': String(row.properties.height) };
    if (row.properties.heightRule) {
      trHeightAttrs['w:hRule'] = row.properties.heightRule;
    }
    children.push({
      tag: 'w:trPr',
      attrs: {},
      children: [{ tag: 'w:trHeight', attrs: trHeightAttrs, children: [] }],
    });
  }

  for (const cell of row.cells) {
    children.push(serializeTableCell(cell));
  }

  return { tag: 'w:tr', attrs: {}, children };
}

function serializeTableCell(cell: TableCell): ParsedNode {
  const children: ParsedNode[] = [];

  if (cell.properties) {
    const tcPrChildren: ParsedNode[] = [];
    if (cell.properties.width) {
      tcPrChildren.push({ tag: 'w:tcW', attrs: { 'w:w': String(cell.properties.width), 'w:type': 'dxa' }, children: [] });
    }
    if (cell.properties.gridSpan && cell.properties.gridSpan > 1) {
      tcPrChildren.push({ tag: 'w:gridSpan', attrs: { 'w:val': String(cell.properties.gridSpan) }, children: [] });
    }
    if (cell.properties.verticalMerge) {
      const vMergeAttrs: Record<string, string> = {};
      if (cell.properties.verticalMerge === 'restart') {
        vMergeAttrs['w:val'] = 'restart';
      }
      tcPrChildren.push({ tag: 'w:vMerge', attrs: vMergeAttrs, children: [] });
    }
    if (cell.properties.verticalAlign) {
      tcPrChildren.push({ tag: 'w:vAlign', attrs: { 'w:val': cell.properties.verticalAlign }, children: [] });
    }
    if (cell.properties.noWrap) {
      tcPrChildren.push({ tag: 'w:noWrap', attrs: {}, children: [] });
    }
    if (cell.properties.borders) {
      const bdrChildren: ParsedNode[] = [];
      const b = cell.properties.borders;
      if (b.top) bdrChildren.push(serializeBorderStyle('w:top', b.top));
      if (b.left) bdrChildren.push(serializeBorderStyle('w:left', b.left));
      if (b.bottom) bdrChildren.push(serializeBorderStyle('w:bottom', b.bottom));
      if (b.right) bdrChildren.push(serializeBorderStyle('w:right', b.right));
      if (bdrChildren.length > 0) {
        tcPrChildren.push({ tag: 'w:tcBorders', attrs: {}, children: bdrChildren });
      }
    }
    if (tcPrChildren.length > 0) {
      children.push({ tag: 'w:tcPr', attrs: {}, children: tcPrChildren });
    }
  }

  for (const block of cell.blocks) {
    if (block.type === 'paragraph') {
      children.push(serializeParagraph(block));
    } else if (block.type === 'table') {
      children.push(serializeTable(block));
    } else if (block.type === 'image') {
      children.push(serializeImageParagraph(block));
    }
  }

  return { tag: 'w:tc', attrs: {}, children };
}
