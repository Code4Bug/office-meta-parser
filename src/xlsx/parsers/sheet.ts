import type { RawDocument, ParsedNode } from '../../core/types.js';
import type {
  Sheet, Cell, MergedCell, FrozenPanes, SheetSelection,
  OutlineGroup, PageSetup, SheetHeaderFooter, SheetComment,
  ExcelTable, TableColumn, SheetProtection, SharedStringEntry,
} from '../types.js';
import { parseCell } from './cell.js';
import { extractSharedStrings } from './strings.js';
import { extractHyperlinks } from './hyperlink.js';
import { extractAutoFilter, extractDataValidations, extractConditionalFormats, extractPrintArea } from './features.js';
import { extractImages } from './image.js';
import { findChild, parseCellRef } from './utils.js';

export type ExtractSheetsResult = { sheets: Sheet[]; authors: string[] };

export function extractSheets(raw: RawDocument): ExtractSheetsResult {
  const sharedStrings = extractSharedStrings(raw);
  const allAuthors: string[] = [];

  const workbookXml = raw.parts.get('xl/workbook.xml');
  if (workbookXml) {
    const sheetsNode = findChild(workbookXml, 'sheets');
    if (!sheetsNode) return { sheets: [], authors: [] };

    const sheets: Sheet[] = [];
    for (const child of sheetsNode.children) {
      if (typeof child === 'string' || child.tag !== 'sheet') continue;

      const name = child.attrs['name'] || '';
      const relId = child.attrs['r:id'] || '';
      const state = parseSheetState(child.attrs['state']);
      const sheetPath = findSheetPath(raw, relId);

      const { sheet, authors } = buildSheet(raw, name, sheetPath, sharedStrings, state, relId);
      sheets.push(sheet);
      if (authors) allAuthors.push(...authors);
    }

    return { sheets, authors: allAuthors };
  }

  // Fallback: scan for worksheet parts directly
  const sheets: Sheet[] = [];
  for (const [path] of raw.parts) {
    if (path.startsWith('xl/worksheets/sheet') && path.endsWith('.xml')) {
      const { sheet, authors } = buildSheet(raw, path.replace('xl/worksheets/', '').replace('.xml', ''), path, sharedStrings);
      sheets.push(sheet);
      if (authors) allAuthors.push(...authors);
    }
  }
  return { sheets, authors: allAuthors };
}

function parseSheetState(val?: string): 'visible' | 'hidden' | 'veryHidden' | undefined {
  if (val === 'hidden') return 'hidden';
  if (val === 'veryHidden') return 'veryHidden';
  return undefined;
}

function buildSheet(
  raw: RawDocument,
  name: string,
  sheetPath: string | undefined,
  sharedStrings: SharedStringEntry[],
  state?: 'visible' | 'hidden' | 'veryHidden',
  _relId?: string,
): { sheet: Sheet; authors: string[] | undefined } {
  const sheet: Sheet = {
    name,
    cells: sheetPath ? extractCells(raw, sheetPath, sharedStrings) : [],
    mergedCells: sheetPath ? extractMergedCells(raw, sheetPath) : [],
    columnWidths: sheetPath ? extractColumnWidths(raw, sheetPath) : [],
    rowHeights: sheetPath ? extractRowHeights(raw, sheetPath) : [],
    hyperlinks: sheetPath ? extractHyperlinks(raw, sheetPath) : [],
    autoFilter: sheetPath ? extractAutoFilter(raw, sheetPath) : undefined,
    dataValidations: sheetPath ? extractDataValidations(raw, sheetPath) : undefined,
    conditionalFormats: sheetPath ? extractConditionalFormats(raw, sheetPath) : undefined,
    printArea: sheetPath ? extractPrintArea(raw, sheetPath) : undefined,
    images: sheetPath ? extractImages(raw, sheetPath) : undefined,
  };

  if (state) sheet.state = state;

  let authors: string[] | undefined;
  if (sheetPath) {
    const wsXml = raw.parts.get(sheetPath);
    if (wsXml) {
      extractSheetViewProps(wsXml, sheet);
      extractSheetPrProps(wsXml, sheet);
      extractSheetFormatPr(wsXml, sheet);
      extractOutlineGroups(wsXml, sheet);
      extractPageSetupProps(wsXml, sheet);
      extractHeaderFooterProps(wsXml, sheet);
      extractSheetProtection(wsXml, sheet);
      const commentsResult = extractComments(raw, sheetPath);
      if (commentsResult) {
        sheet.comments = commentsResult.comments;
        authors = commentsResult.authors;
      }
      sheet.tables = extractTables(raw, sheetPath);
    }
  }

  return { sheet, authors };
}

function findSheetPath(raw: RawDocument, relId: string): string | undefined {
  const wbRels = raw.rels.get('xl/_rels/workbook.xml.rels');
  if (!wbRels) return undefined;

  const rel = wbRels.find(r => r.id === relId);
  return rel ? 'xl/' + rel.target : undefined;
}

export function extractCells(raw: RawDocument, sheetPath: string, sharedStrings: SharedStringEntry[]): Cell[][] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  const sheetData = findChild(worksheetXml, 'sheetData');
  if (!sheetData) return [];

  const rows: Cell[][] = [];
  for (const rowNode of sheetData.children) {
    if (typeof rowNode === 'string' || rowNode.tag !== 'row') continue;

    const cells: Cell[] = [];
    for (const cellNode of rowNode.children) {
      if (typeof cellNode === 'string' || cellNode.tag !== 'c') continue;
      cells.push(parseCell(cellNode, sharedStrings));
    }
    rows.push(cells);
  }

  return rows;
}

export function extractMergedCells(raw: RawDocument, sheetPath: string): MergedCell[] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  const mergeCellsNode = findChild(worksheetXml, 'mergeCells');
  if (!mergeCellsNode) return [];

  const merged: MergedCell[] = [];
  for (const child of mergeCellsNode.children) {
    if (typeof child === 'string' || child.tag !== 'mergeCell') continue;
    const ref = child.attrs['ref'];
    if (!ref) continue;

    const parts = ref.split(':');
    if (parts.length !== 2) continue;

    const start = parseCellRef(parts[0]);
    const end = parseCellRef(parts[1]);
    if (start && end) {
      merged.push({
        startRow: start.row,
        startCol: start.col,
        endRow: end.row,
        endCol: end.col,
      });
    }
  }

  return merged;
}

export function extractColumnWidths(raw: RawDocument, sheetPath: string): number[] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  const colsNode = findChild(worksheetXml, 'cols');
  if (!colsNode) return [];

  const widths: number[] = [];
  for (const child of colsNode.children) {
    if (typeof child === 'string' || child.tag !== 'col') continue;
    const width = child.attrs['width'];
    if (width) widths.push(parseFloat(width));
  }

  return widths;
}

export function extractRowHeights(raw: RawDocument, sheetPath: string): number[] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  const sheetData = findChild(worksheetXml, 'sheetData');
  if (!sheetData) return [];

  const heights: number[] = [];
  for (const rowNode of sheetData.children) {
    if (typeof rowNode === 'string' || rowNode.tag !== 'row') continue;
    const ht = rowNode.attrs['ht'];
    if (ht) heights.push(parseFloat(ht));
  }

  return heights;
}

function extractSheetViewProps(wsXml: ParsedNode, sheet: Sheet): void {
  const sheetViews = findChild(wsXml, 'sheetViews');
  if (!sheetViews) return;
  const sheetView = findChild(sheetViews, 'sheetView');
  if (!sheetView) return;

  if (sheetView.attrs['zoomScale']) {
    sheet.zoomScale = parseInt(sheetView.attrs['zoomScale'], 10);
  }

  const selections: SheetSelection[] = [];
  for (const child of sheetView.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'selection') {
      const sel: SheetSelection = {};
      if (child.attrs['pane']) sel.pane = child.attrs['pane'];
      if (child.attrs['activeCell']) sel.activeCell = child.attrs['activeCell'];
      if (child.attrs['sqref']) sel.sqref = child.attrs['sqref'];
      selections.push(sel);
    }
  }
  if (selections.length > 0) sheet.selections = selections;

  if (selections.length > 0 && selections[0].activeCell) {
    sheet.activeCell = selections[0].activeCell;
  }

  const pane = findChild(sheetView, 'pane');
  if (pane) {
    const fp: FrozenPanes = {};
    if (pane.attrs['xSplit']) fp.xSplit = parseInt(pane.attrs['xSplit'], 10);
    if (pane.attrs['ySplit']) fp.ySplit = parseInt(pane.attrs['ySplit'], 10);
    if (pane.attrs['topLeftCell']) fp.topLeftCell = pane.attrs['topLeftCell'];
    if (fp.xSplit || fp.ySplit) sheet.frozenPanes = fp;
  }
}

function extractSheetPrProps(wsXml: ParsedNode, sheet: Sheet): void {
  const sheetPr = findChild(wsXml, 'sheetPr');
  if (!sheetPr) return;

  const tabColor = findChild(sheetPr, 'tabColor');
  if (tabColor) {
    sheet.tabColor = parseColorValue(tabColor);
  }
}

function parseColorValue(node: ParsedNode): string {
  if (node.attrs['rgb']) return node.attrs['rgb'];
  if (node.attrs['indexed']) return `indexed:${node.attrs['indexed']}`;
  if (node.attrs['theme']) return `theme:${node.attrs['theme']}`;
  if (node.attrs['val']) return node.attrs['val'];
  return '';
}

function extractSheetFormatPr(wsXml: ParsedNode, sheet: Sheet): void {
  const formatPr = findChild(wsXml, 'sheetFormatPr');
  if (!formatPr) return;

  if (formatPr.attrs['defaultRowHeight']) {
    sheet.defaultRowHeight = parseFloat(formatPr.attrs['defaultRowHeight']);
  }
  if (formatPr.attrs['defaultColWidth']) {
    sheet.defaultColWidth = parseFloat(formatPr.attrs['defaultColWidth']);
  }
}

function extractOutlineGroups(wsXml: ParsedNode, sheet: Sheet): void {
  const cols = findChild(wsXml, 'cols');
  if (cols) {
    const colGroups: OutlineGroup[] = [];
    for (const child of cols.children) {
      if (typeof child === 'string' || child.tag !== 'col') continue;
      const outlineLevel = parseInt(child.attrs['outlineLevel'] || '0', 10);
      if (outlineLevel > 0) {
        colGroups.push({
          level: outlineLevel,
          collapsed: child.attrs['collapsed'] === '1',
        });
      }
    }
    if (colGroups.length > 0) sheet.colGroups = colGroups;
  }

  const sheetData = findChild(wsXml, 'sheetData');
  if (sheetData) {
    const rowGroups: OutlineGroup[] = [];
    for (const child of sheetData.children) {
      if (typeof child === 'string' || child.tag !== 'row') continue;
      const outlineLevel = parseInt(child.attrs['outlineLevel'] || '0', 10);
      if (outlineLevel > 0) {
        rowGroups.push({
          level: outlineLevel,
          collapsed: child.attrs['collapsed'] === '1',
        });
      }
    }
    if (rowGroups.length > 0) sheet.rowGroups = rowGroups;
  }
}

function extractPageSetupProps(wsXml: ParsedNode, sheet: Sheet): void {
  const pageSetup = findChild(wsXml, 'pageSetup');
  if (!pageSetup) return;

  const ps: PageSetup = {};
  if (pageSetup.attrs['orientation']) {
    const o = pageSetup.attrs['orientation'];
    if (o === 'portrait' || o === 'landscape') ps.orientation = o;
  }
  if (pageSetup.attrs['paperSize']) ps.paperSize = parseInt(pageSetup.attrs['paperSize'], 10);
  if (pageSetup.attrs['scale']) ps.scale = parseInt(pageSetup.attrs['scale'], 10);
  if (pageSetup.attrs['fitToWidth']) ps.fitToWidth = parseInt(pageSetup.attrs['fitToWidth'], 10);
  if (pageSetup.attrs['fitToHeight']) ps.fitToHeight = parseInt(pageSetup.attrs['fitToHeight'], 10);
  if (pageSetup.attrs['firstPageNumber']) ps.firstPageNumber = parseInt(pageSetup.attrs['firstPageNumber'], 10);
  if (pageSetup.attrs['useFirstPageNumber'] === '1') ps.useFirstPageNumber = true;
  if (pageSetup.attrs['horizontalDpi']) ps.horizontalDpi = parseInt(pageSetup.attrs['horizontalDpi'], 10);
  if (pageSetup.attrs['verticalDpi']) ps.verticalDpi = parseInt(pageSetup.attrs['verticalDpi'], 10);

  sheet.pageSetup = ps;
}

function extractHeaderFooterProps(wsXml: ParsedNode, sheet: Sheet): void {
  const hf = findChild(wsXml, 'headerFooter');
  if (!hf) return;

  const headerFooter: SheetHeaderFooter = {};
  if (hf.attrs['differentFirst'] === '1') headerFooter.differentFirst = true;
  if (hf.attrs['differentOddEven'] === '1') headerFooter.differentOddEven = true;

  for (const child of hf.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'oddHeader': headerFooter.oddHeader = getDirectText(child); break;
      case 'oddFooter': headerFooter.oddFooter = getDirectText(child); break;
      case 'evenHeader': headerFooter.evenHeader = getDirectText(child); break;
      case 'evenFooter': headerFooter.evenFooter = getDirectText(child); break;
      case 'firstHeader': headerFooter.firstHeader = getDirectText(child); break;
      case 'firstFooter': headerFooter.firstFooter = getDirectText(child); break;
    }
  }

  sheet.headerFooter = headerFooter;
}

function getDirectText(node: ParsedNode): string {
  for (const child of node.children) {
    if (typeof child === 'string') return child;
  }
  return '';
}

export type ExtractCommentsResult = { authors: string[]; comments: SheetComment[] };

function extractComments(raw: RawDocument, sheetPath: string): ExtractCommentsResult | undefined {
  const sheetName = sheetPath.substring(sheetPath.lastIndexOf('/') + 1).replace('.xml', '');
  const commentsPath = `xl/comments${sheetName.replace('sheet', '')}.xml`;

  const commentsXml = raw.parts.get(commentsPath);
  if (!commentsXml) return undefined;

  const authors: string[] = [];
  const authorsNode = findChild(commentsXml, 'authors');
  if (authorsNode) {
    for (const child of authorsNode.children) {
      if (typeof child === 'string' || child.tag !== 'author') continue;
      for (const t of child.children) {
        if (typeof t === 'string') { authors.push(t); break; }
      }
    }
  }

  const commentsNode = findChild(commentsXml, 'commentList');
  if (!commentsNode) return undefined;

  const comments: SheetComment[] = [];
  for (const child of commentsNode.children) {
    if (typeof child === 'string' || child.tag !== 'comment') continue;
    const ref = child.attrs['ref'];
    const authorId = parseInt(child.attrs['authorId'] || '0', 10);
    const textNode = findChild(child, 'text');
    if (!textNode) continue;

    let text = '';
    for (const t of textNode.children) {
      if (typeof t === 'string') text += t;
      else if (t.tag === 't') text += getDirectText(t);
    }

    if (ref) {
      comments.push({ ref, authorId, text });
    }
  }

  return comments.length > 0 ? { authors, comments } : undefined;
}

function extractTables(raw: RawDocument, sheetPath: string): ExcelTable[] | undefined {
  const sheetDir = sheetPath.substring(0, sheetPath.lastIndexOf('/'));
  const sheetBase = sheetPath.substring(sheetPath.lastIndexOf('/') + 1);
  const relsPath = `${sheetDir}/_rels/${sheetBase}.rels`;
  const sheetRels = raw.rels.get(relsPath);

  if (!sheetRels) return undefined;

  const tables: ExcelTable[] = [];
  for (const rel of sheetRels) {
    if (!rel.type.includes('table')) continue;
    // Resolve relative path: ../tables/table1.xml relative to xl/worksheets/
    const tablePath = resolveRelative(sheetDir + '/', rel.target);
    const tableXml = raw.parts.get(tablePath);
    if (!tableXml) continue;

    const table = parseTable(tableXml);
    if (table) tables.push(table);
  }

  return tables.length > 0 ? tables : undefined;
}

function parseTable(node: ParsedNode): ExcelTable | undefined {
  const id = parseInt(node.attrs['id'] || '0', 10);
  const name = node.attrs['name'] || '';
  const displayName = node.attrs['displayName'] || name;
  const ref = node.attrs['ref'] || '';
  if (!ref) return undefined;

  const table: ExcelTable = {
    id,
    name,
    displayName,
    ref,
    columns: [],
  };

  if (node.attrs['headerRowCount']) table.headerRowCount = parseInt(node.attrs['headerRowCount'], 10);
  if (node.attrs['totalsRowCount']) table.totalsRowCount = parseInt(node.attrs['totalsRowCount'], 10);

  const tableColumns = findChild(node, 'tableColumns');
  if (tableColumns) {
    for (const child of tableColumns.children) {
      if (typeof child === 'string' || child.tag !== 'tableColumn') continue;
      const col: TableColumn = {
        id: parseInt(child.attrs['id'] || '0', 10),
        name: child.attrs['name'] || '',
      };
      if (child.attrs['totalsRowLabel']) col.totalsRowLabel = child.attrs['totalsRowLabel'];
      if (child.attrs['totalsRowFunction']) col.totalsRowFunction = child.attrs['totalsRowFunction'];
      table.columns.push(col);
    }
  }

  return table;
}

function extractSheetProtection(wsXml: ParsedNode, sheet: Sheet): void {
  const sp = findChild(wsXml, 'sheetProtection');
  if (!sp) return;

  const prot: SheetProtection = {};
  if (sp.attrs['sheet'] === '1') prot.enabled = true;
  if (sp.attrs['password']) prot.password = sp.attrs['password'];
  if (sp.attrs['selectLockedCells']) prot.selectLockedCells = sp.attrs['selectLockedCells'] !== '0';
  if (sp.attrs['selectUnlockedCells']) prot.selectUnlockedCells = sp.attrs['selectUnlockedCells'] !== '0';
  if (sp.attrs['insertRows']) prot.insertRows = sp.attrs['insertRows'] === '1';
  if (sp.attrs['deleteRows']) prot.deleteRows = sp.attrs['deleteRows'] === '1';
  if (sp.attrs['formatCells']) prot.formatCells = sp.attrs['formatCells'] === '1';
  if (sp.attrs['sort']) prot.sort = sp.attrs['sort'] === '1';
  if (sp.attrs['autoFilter']) prot.autoFilter = sp.attrs['autoFilter'] === '1';
  sheet.protection = prot;
}

function resolveRelative(base: string, relative: string): string {
  const baseParts = base.split('/').filter(Boolean);
  const relParts = relative.split('/').filter(Boolean);
  const result = [...baseParts];
  for (const part of relParts) {
    if (part === '..') result.pop();
    else if (part !== '.') result.push(part);
  }
  return result.join('/');
}
