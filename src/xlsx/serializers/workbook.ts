import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { XlsxWorkbook } from '../types.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function serializeWorkbook(wb: XlsxWorkbook): string {
  const sheetsChildren = wb.sheets.map((sheet, i) => ({
    tag: 'sheet',
    attrs: {
      name: sheet.name,
      sheetId: String(i + 1),
      'r:id': `rId${i + 1}`,
      ...(sheet.state ? { state: sheet.state } : {}),
    },
    children: [],
  }));

  const children: ParsedNode[] = [
    { tag: 'sheets', attrs: {}, children: sheetsChildren },
  ];

  // Collect definedNames including printTitles from sheets
  const allDefinedNames = [...(wb.definedNames || [])];
  wb.sheets.forEach((sheet, i) => {
    if (sheet.printTitles) {
      const already = allDefinedNames.find(
        dn => dn.name === '_xlnm.Print_Titles' && dn.localSheetId === i
      );
      if (!already) {
        allDefinedNames.push({ name: '_xlnm.Print_Titles', formula: sheet.printTitles, localSheetId: i });
      }
    }
  });

  if (allDefinedNames.length > 0) {
    const dns = allDefinedNames.map(dn => {
      const attrs: Record<string, string> = { name: dn.name };
      if (dn.localSheetId !== undefined) attrs.localSheetId = String(dn.localSheetId);
      if (dn.hidden) attrs.hidden = '1';
      return { tag: 'definedName', attrs, children: [dn.formula] };
    });
    children.push({ tag: 'definedNames', attrs: {}, children: dns });
  }

  const root: ParsedNode = {
    tag: 'workbook',
    attrs: { xmlns: MAIN_NS, 'xmlns:r': R_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
