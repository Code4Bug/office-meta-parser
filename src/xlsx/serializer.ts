import type { XlsxWorkbook } from './types.js';
import { serializeWorkbook } from './serializers/workbook.js';
import { serializeWorksheet } from './serializers/worksheet.js';
import { serializeDrawing } from './serializers/drawing.js';
import { serializeSharedStrings } from './serializers/shared-strings.js';
import { serializeCellStyles } from './serializers/styles.js';
import { serializeWorkbookRels, serializeSheetRels, serializeContentTypes, serializeCommentsXml } from './serializers/rels.js';

export interface XlsxXmlParts {
  workbook: string;
  worksheets: string[];
  sharedStrings: string;
  styles: string;
  rels: string;
  sheetRels: string[];
  contentTypes: string;
  drawings: string[];
  comments: { path: string; xml: string }[];
}

export function semanticToXml(wb: XlsxWorkbook): XlsxXmlParts {
  const comments: { path: string; xml: string }[] = [];
  wb.sheets.forEach((sheet, i) => {
    if (sheet.comments && sheet.comments.length > 0) {
      comments.push({
        path: `xl/comments${i + 1}.xml`,
        xml: serializeCommentsXml(i, 'Author', sheet.comments),
      });
    }
  });

  return {
    workbook: serializeWorkbook(wb),
    worksheets: wb.sheets.map((sheet, i) => serializeWorksheet(sheet, i, wb.sharedStrings)),
    sharedStrings: serializeSharedStrings(wb.sharedStrings),
    styles: serializeCellStyles(wb.styles),
    rels: serializeWorkbookRels(wb.sheets),
    sheetRels: wb.sheets.map((sheet, i) => serializeSheetRels(sheet, i)),
    contentTypes: serializeContentTypes(wb.sheets, wb.extraContentTypes),
    drawings: wb.sheets.map(sheet => sheet.images && sheet.images.length > 0 ? serializeDrawing(sheet.images) : ''),
    comments,
  };
}

export { serializeCellStyles } from './serializers/styles.js';
