import type { ParsedNode } from '../../core/types.js';
import type {
  Table, TableRow, TableCell, DocxBlock,
  TableProperties, TableRowProperties, TableCellProperties,
  BorderStyle, TableBorders, TableCellBorders,
} from '../types.js';
import { parseParagraph } from './paragraph.js';
import { findChild } from './utils.js';

export function parseTable(node: ParsedNode): Table {
  const rows: TableRow[] = [];
  let properties: TableProperties | undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:tblPr') {
      properties = parseTableProperties(child);
    } else if (child.tag === 'w:tr') {
      rows.push(parseTableRow(child));
    }
  }

  const result: Table = { type: 'table', rows };
  if (properties) result.properties = properties;
  return result;
}

function parseTableProperties(node: ParsedNode): TableProperties {
  const props: TableProperties = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:tblStyle') {
      if (child.attrs['w:val']) props.style = child.attrs['w:val'];
    } else if (child.tag === 'w:tblW') {
      if (child.attrs['w:w']) props.width = parseInt(child.attrs['w:w'], 10);
      if (child.attrs['w:type']) props.widthType = child.attrs['w:type'];
    } else if (child.tag === 'w:tblInd') {
      if (child.attrs['w:w']) props.indent = parseInt(child.attrs['w:w'], 10);
    } else if (child.tag === 'w:tblLayout') {
      if (child.attrs['w:type']) {
        props.layout = child.attrs['w:type'] as 'autofit' | 'fixed';
      }
    } else if (child.tag === 'w:tblCellMar') {
      const top = child.children.find(c => typeof c !== 'string' && c.tag === 'w:top');
      const left = child.children.find(c => typeof c !== 'string' && c.tag === 'w:left');
      const bottom = child.children.find(c => typeof c !== 'string' && c.tag === 'w:bottom');
      const right = child.children.find(c => typeof c !== 'string' && c.tag === 'w:right');
      if (top && typeof top !== 'string' && top.attrs['w:w']) props.cellMarginTop = parseInt(top.attrs['w:w'], 10);
      if (left && typeof left !== 'string' && left.attrs['w:w']) props.cellMarginLeft = parseInt(left.attrs['w:w'], 10);
      if (bottom && typeof bottom !== 'string' && bottom.attrs['w:w']) props.cellMarginBottom = parseInt(bottom.attrs['w:w'], 10);
      if (right && typeof right !== 'string' && right.attrs['w:w']) props.cellMarginRight = parseInt(right.attrs['w:w'], 10);
    } else if (child.tag === 'w:tblGrid') {
      props.gridColumns = [];
      for (const gridCol of child.children) {
        if (typeof gridCol === 'string') continue;
        if (gridCol.tag === 'w:gridCol' && gridCol.attrs['w:w']) {
          props.gridColumns.push(parseInt(gridCol.attrs['w:w'], 10));
        }
      }
    } else if (child.tag === 'w:tblBorders') {
      props.borders = parseTableBorders(child);
    }
  }

  return props;
}

function parseTableBorders(node: ParsedNode): TableBorders {
  const borders: TableBorders = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    const border = parseBorderStyle(child);
    switch (child.tag) {
      case 'w:top': borders.top = border; break;
      case 'w:bottom': borders.bottom = border; break;
      case 'w:left': borders.left = border; break;
      case 'w:right': borders.right = border; break;
      case 'w:insideH': borders.insideHorizontal = border; break;
      case 'w:insideV': borders.insideVertical = border; break;
    }
  }

  return borders;
}

function parseBorderStyle(node: ParsedNode): BorderStyle {
  const style: BorderStyle = { style: node.attrs['w:val'] || 'none' };
  if (node.attrs['w:sz']) style.size = parseInt(node.attrs['w:sz'], 10);
  if (node.attrs['w:color']) style.color = node.attrs['w:color'];
  return style;
}

function parseTableRow(node: ParsedNode): TableRow {
  const cells: TableCell[] = [];
  let properties: TableRowProperties | undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:trPr') {
      properties = {};
      const trHeight = findChild(child, 'w:trHeight');
      if (trHeight?.attrs['w:val']) {
        properties.height = parseInt(trHeight.attrs['w:val'], 10);
      }
      if (trHeight?.attrs['w:hRule']) {
        properties.heightRule = trHeight.attrs['w:hRule'] as 'auto' | 'exact' | 'atLeast';
      }
    } else if (child.tag === 'w:tc') {
      cells.push(parseTableCell(child));
    }
  }

  const result: TableRow = { cells };
  if (properties) result.properties = properties;
  return result;
}

function parseTableCell(node: ParsedNode): TableCell {
  const blocks: DocxBlock[] = [];
  let properties: TableCellProperties | undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:tcPr') {
      properties = parseTableCellProperties(child);
    } else if (child.tag === 'w:p') {
      blocks.push(parseParagraph(child));
    } else if (child.tag === 'w:tbl') {
      blocks.push(parseTable(child));
    }
  }

  const result: TableCell = { blocks };
  if (properties) result.properties = properties;
  return result;
}

function parseTableCellProperties(node: ParsedNode): TableCellProperties {
  const props: TableCellProperties = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:tcW') {
      if (child.attrs['w:w']) props.width = parseInt(child.attrs['w:w'], 10);
    } else if (child.tag === 'w:vAlign') {
      if (child.attrs['w:val']) {
        const val = child.attrs['w:val'];
        if (val === 'top' || val === 'center' || val === 'bottom') {
          props.verticalAlign = val;
        }
      }
    } else if (child.tag === 'w:vMerge') {
      props.verticalMerge = child.attrs['w:val'] === 'restart' ? 'restart' : 'continue';
    } else if (child.tag === 'w:gridSpan') {
      if (child.attrs['w:val']) {
        const span = parseInt(child.attrs['w:val'], 10);
        props.gridSpan = span;
        if (span > 1) props.horizontalMerge = 'restart';
      }
    } else if (child.tag === 'w:tcBorders') {
      props.borders = parseTableCellBorders(child);
    } else if (child.tag === 'w:noWrap') {
      props.noWrap = true;
    }
  }

  return props;
}

function parseTableCellBorders(node: ParsedNode): TableCellBorders {
  const borders: TableCellBorders = {};
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    const border = parseBorderStyle(child);
    switch (child.tag) {
      case 'w:top': borders.top = border; break;
      case 'w:bottom': borders.bottom = border; break;
      case 'w:left': borders.left = border; break;
      case 'w:right': borders.right = border; break;
    }
  }
  return borders;
}
