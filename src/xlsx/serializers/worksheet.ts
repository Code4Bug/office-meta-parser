import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { Sheet, AutoFilter, DataValidation, ConditionalFormat, PrintArea, SharedStringEntry } from '../types.js';
import { serializeCell, getColLetter } from './cell.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function serializeWorksheet(sheet: Sheet, _index: number, sharedStrings: SharedStringEntry[]): string {
  const children: ParsedNode[] = [];

  // Sheet properties (tab color)
  if (sheet.tabColor) {
    children.push({
      tag: 'sheetPr',
      attrs: {},
      children: [serializeColor(sheet.tabColor, 'tabColor')],
    });
  }

  // Sheet view (frozen panes, zoom, selections)
  children.push(serializeSheetView(sheet));

  // Sheet format (default row/col dimensions)
  if (sheet.defaultRowHeight !== undefined || sheet.defaultColWidth !== undefined) {
    const attrs: Record<string, string> = {};
    if (sheet.defaultRowHeight !== undefined) attrs.defaultRowHeight = String(sheet.defaultRowHeight);
    if (sheet.defaultColWidth !== undefined) attrs.defaultColWidth = String(sheet.defaultColWidth);
    children.push({ tag: 'sheetFormatPr', attrs, children: [] });
  }

  // Column widths
  if (sheet.columnWidths.length > 0 || (sheet.colGroups && sheet.colGroups.length > 0)) {
    const colsChildren: ParsedNode[] = [];

    if (sheet.columnWidths.length > 0) {
      sheet.columnWidths.forEach((width, i) => {
        const attrs: Record<string, string> = {
          min: String(i + 1), max: String(i + 1), width: String(width), customWidth: '1',
        };
        if (sheet.colGroups && sheet.colGroups[i]) {
          attrs.outlineLevel = String(sheet.colGroups[i].level);
          if (sheet.colGroups[i].collapsed) attrs.collapsed = '1';
        }
        colsChildren.push({ tag: 'col', attrs, children: [] });
      });
    }

    children.push({ tag: 'cols', attrs: {}, children: colsChildren });
  }

  // Sheet data
  const rows: ParsedNode[] = [];
  sheet.cells.forEach((row, rowIdx) => {
    const cells: ParsedNode[] = row.map((cell, colIdx) => {
      return serializeCell(cell, rowIdx, colIdx, sharedStrings);
    });

    const rowAttrs: Record<string, string> = { r: String(rowIdx + 1) };
    if (sheet.rowHeights[rowIdx]) {
      rowAttrs.ht = String(sheet.rowHeights[rowIdx]);
      rowAttrs.customHeight = '1';
    }
    if (sheet.rowGroups && sheet.rowGroups[rowIdx]) {
      rowAttrs.outlineLevel = String(sheet.rowGroups[rowIdx].level);
      if (sheet.rowGroups[rowIdx].collapsed) rowAttrs.collapsed = '1';
    }

    rows.push({ tag: 'row', attrs: rowAttrs, children: cells });
  });
  children.push({ tag: 'sheetData', attrs: {}, children: rows });

  // Merged cells
  if (sheet.mergedCells.length > 0) {
    const mergeCells = sheet.mergedCells.map(mc => ({
      tag: 'mergeCell',
      attrs: { ref: `${getColLetter(mc.startCol)}${mc.startRow + 1}:${getColLetter(mc.endCol)}${mc.endRow + 1}` },
      children: [],
    }));
    children.push({ tag: 'mergeCells', attrs: { count: String(mergeCells.length) }, children: mergeCells });
  }

  // Sheet protection
  if (sheet.protection) {
    const p = sheet.protection;
    const protAttrs: Record<string, string> = {};
    if (p.enabled) protAttrs.sheet = '1';
    if (p.password) protAttrs.password = p.password;
    if (p.selectLockedCells !== undefined) protAttrs.selectLockedCells = p.selectLockedCells ? '1' : '0';
    if (p.selectUnlockedCells !== undefined) protAttrs.selectUnlockedCells = p.selectUnlockedCells ? '1' : '0';
    if (p.insertRows !== undefined) protAttrs.insertRows = p.insertRows ? '1' : '0';
    if (p.deleteRows !== undefined) protAttrs.deleteRows = p.deleteRows ? '1' : '0';
    if (p.formatCells !== undefined) protAttrs.formatCells = p.formatCells ? '1' : '0';
    if (p.sort !== undefined) protAttrs.sort = p.sort ? '1' : '0';
    if (p.autoFilter !== undefined) protAttrs.autoFilter = p.autoFilter ? '1' : '0';
    children.push({ tag: 'sheetProtection', attrs: protAttrs, children: [] });
  }

  // Auto filter
  if (sheet.autoFilter) {
    children.push(serializeAutoFilter(sheet.autoFilter));
  }

  // Data validations
  if (sheet.dataValidations && sheet.dataValidations.length > 0) {
    children.push(serializeDataValidations(sheet.dataValidations));
  }

  // Conditional formats
  if (sheet.conditionalFormats && sheet.conditionalFormats.length > 0) {
    for (const cf of sheet.conditionalFormats) {
      children.push(serializeConditionalFormat(cf));
    }
  }

  // Page setup
  if (sheet.pageSetup) {
    const ps = sheet.pageSetup;
    const attrs: Record<string, string> = {};
    if (ps.orientation) attrs.orientation = ps.orientation;
    if (ps.paperSize !== undefined) attrs.paperSize = String(ps.paperSize);
    if (ps.scale !== undefined) attrs.scale = String(ps.scale);
    if (ps.fitToWidth !== undefined) attrs.fitToWidth = String(ps.fitToWidth);
    if (ps.fitToHeight !== undefined) attrs.fitToHeight = String(ps.fitToHeight);
    if (ps.firstPageNumber !== undefined) attrs.firstPageNumber = String(ps.firstPageNumber);
    if (ps.useFirstPageNumber) attrs.useFirstPageNumber = '1';
    if (ps.horizontalDpi !== undefined) attrs.horizontalDpi = String(ps.horizontalDpi);
    if (ps.verticalDpi !== undefined) attrs.verticalDpi = String(ps.verticalDpi);
    children.push({ tag: 'pageSetup', attrs, children: [] });
  }

  // Header/footer
  if (sheet.headerFooter) {
    children.push(serializeHeaderFooter(sheet.headerFooter));
  }

  // Print area (existing)
  if (sheet.printArea) {
    children.push(...serializePrintArea(sheet.printArea));
  }

  // Tables
  if (sheet.tables && sheet.tables.length > 0) {
    const tableParts = sheet.tables.map((_, i) => ({
      tag: 'tablePart',
      attrs: { 'r:id': `rId${i + 1}` },
      children: [],
    }));
    children.push({ tag: 'tableParts', attrs: { count: String(tableParts.length) }, children: tableParts });
  }

  // Drawing reference
  if (sheet.images && sheet.images.length > 0) {
    children.push({ tag: 'drawing', attrs: { 'r:id': 'rId1' }, children: [] });
  }

  const root: ParsedNode = {
    tag: 'worksheet',
    attrs: { xmlns: MAIN_NS, 'xmlns:r': R_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeAutoFilter(autoFilter: AutoFilter): ParsedNode {
  const columns = autoFilter.columns.map(col => {
    const filterChildren: ParsedNode[] = [];

    if (col.filters && col.filters.length > 0) {
      const filters = col.filters.map(val => ({
        tag: 'filter',
        attrs: { val },
        children: [],
      }));
      filterChildren.push({
        tag: 'filters',
        attrs: {},
        children: filters,
      });
    }

    if (col.customFilter) {
      filterChildren.push({
        tag: 'customFilters',
        attrs: {},
        children: [{
          tag: 'customFilter',
          attrs: { operator: col.customFilter.operator, val: col.customFilter.value },
          children: [],
        }],
      });
    }

    return {
      tag: 'filterColumn',
      attrs: { colId: String(col.colId) },
      children: filterChildren,
    };
  });

  return {
    tag: 'autoFilter',
    attrs: { ref: autoFilter.ref },
    children: columns,
  };
}

function serializeDataValidations(validations: DataValidation[]): ParsedNode {
  const children = validations.map(v => {
    const attrs: Record<string, string> = {
      type: v.type,
      sqref: v.sqref,
    };

    if (v.operator) attrs.operator = v.operator;
    if (v.allowBlank !== undefined) attrs.allowBlank = v.allowBlank ? '1' : '0';
    if (v.showErrorMessage !== undefined) attrs.showErrorMessage = v.showErrorMessage ? '1' : '0';
    if (v.errorTitle) attrs.errorTitle = v.errorTitle;
    if (v.error) attrs.error = v.error;

    const formulaChildren: ParsedNode[] = [];
    if (v.formula1) {
      formulaChildren.push({ tag: 'formula1', attrs: {}, children: [v.formula1] });
    }
    if (v.formula2) {
      formulaChildren.push({ tag: 'formula2', attrs: {}, children: [v.formula2] });
    }

    return {
      tag: 'dataValidation',
      attrs,
      children: formulaChildren,
    };
  });

  return {
    tag: 'dataValidations',
    attrs: { count: String(validations.length) },
    children,
  };
}

function serializeConditionalFormat(cf: ConditionalFormat): ParsedNode {
  const rules = cf.rules.map(rule => {
    const attrs: Record<string, string> = {
      type: rule.type,
      priority: String(rule.priority),
    };

    const ruleChildren: ParsedNode[] = [];

    if (rule.formula && rule.formula.length > 0) {
      ruleChildren.push({ tag: 'formula', attrs: {}, children: [rule.formula[0]] });
    }

    if (rule.colorScale) {
      const colors = rule.colorScale.colors.map(color => ({
        tag: 'color',
        attrs: { rgb: color },
        children: [],
      }));
      ruleChildren.push({ tag: 'colorScale', attrs: {}, children: colors });
    }

    return {
      tag: 'cfRule',
      attrs,
      children: ruleChildren,
    };
  });

  return {
    tag: 'conditionalFormatting',
    attrs: { sqref: cf.sqref },
    children: rules,
  };
}

function serializePrintArea(printArea: PrintArea): ParsedNode[] {
  const result: ParsedNode[] = [];

  if (printArea.fitToWidth !== undefined || printArea.fitToHeight !== undefined) {
    const attrs: Record<string, string> = {};
    if (printArea.fitToWidth !== undefined) attrs.fitToWidth = String(printArea.fitToWidth);
    if (printArea.fitToHeight !== undefined) attrs.fitToHeight = String(printArea.fitToHeight);
    result.push({ tag: 'printOptions', attrs, children: [] });
  }

  if (printArea.pageMargins) {
    const m = printArea.pageMargins;
    result.push({
      tag: 'pageMargins',
      attrs: {
        top: String(m.top),
        right: String(m.right),
        bottom: String(m.bottom),
        left: String(m.left),
        header: String(m.header),
        footer: String(m.footer),
      },
      children: [],
    });
  }

  return result;
}

function serializeSheetView(sheet: Sheet): ParsedNode {
  const viewAttrs: Record<string, string> = { workbookViewId: '0' };
  if (sheet.zoomScale !== undefined) viewAttrs.zoomScale = String(sheet.zoomScale);

  const viewChildren: ParsedNode[] = [];

  if (sheet.frozenPanes) {
    const paneAttrs: Record<string, string> = {};
    if (sheet.frozenPanes.xSplit !== undefined) paneAttrs.xSplit = String(sheet.frozenPanes.xSplit);
    if (sheet.frozenPanes.ySplit !== undefined) paneAttrs.ySplit = String(sheet.frozenPanes.ySplit);
    if (sheet.frozenPanes.topLeftCell) paneAttrs.topLeftCell = sheet.frozenPanes.topLeftCell;
    paneAttrs.activePane = 'bottomRight';
    paneAttrs.state = 'frozen';
    viewChildren.push({ tag: 'pane', attrs: paneAttrs, children: [] });
  }

  if (sheet.selections && sheet.selections.length > 0) {
    for (const sel of sheet.selections) {
      const attrs: Record<string, string> = {};
      if (sel.pane) attrs.pane = sel.pane;
      if (sel.activeCell) attrs.activeCell = sel.activeCell;
      if (sel.sqref) attrs.sqref = sel.sqref;
      viewChildren.push({ tag: 'selection', attrs, children: [] });
    }
  } else if (sheet.activeCell) {
    viewChildren.push({
      tag: 'selection',
      attrs: { activeCell: sheet.activeCell, sqref: sheet.activeCell },
      children: [],
    });
  }

  return {
    tag: 'sheetViews',
    attrs: {},
    children: [{ tag: 'sheetView', attrs: viewAttrs, children: viewChildren }],
  };
}

function serializeColor(color: string, tag: string): ParsedNode {
  if (color.startsWith('theme:')) {
    return { tag, attrs: { theme: color.substring(6) }, children: [] };
  }
  if (color.startsWith('indexed:')) {
    return { tag, attrs: { indexed: color.substring(8) }, children: [] };
  }
  return { tag, attrs: { rgb: color }, children: [] };
}

function serializeHeaderFooter(hf: import('../types.js').SheetHeaderFooter): ParsedNode {
  const attrs: Record<string, string> = {};
  if (hf.differentFirst) attrs.differentFirst = '1';
  if (hf.differentOddEven) attrs.differentOddEven = '1';

  const children: ParsedNode[] = [];
  if (hf.oddHeader) children.push({ tag: 'oddHeader', attrs: {}, children: [hf.oddHeader] });
  if (hf.oddFooter) children.push({ tag: 'oddFooter', attrs: {}, children: [hf.oddFooter] });
  if (hf.evenHeader) children.push({ tag: 'evenHeader', attrs: {}, children: [hf.evenHeader] });
  if (hf.evenFooter) children.push({ tag: 'evenFooter', attrs: {}, children: [hf.evenFooter] });
  if (hf.firstHeader) children.push({ tag: 'firstHeader', attrs: {}, children: [hf.firstHeader] });
  if (hf.firstFooter) children.push({ tag: 'firstFooter', attrs: {}, children: [hf.firstFooter] });

  return { tag: 'headerFooter', attrs, children };
}
