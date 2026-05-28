import type { ValidationIssue } from '../core/validate.js';
import type { XlsxWorkbook, Sheet, Cell } from './types.js';

export function validateXlsx(wb: XlsxWorkbook): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  // ═══════════════ 必须项 (error) ═══════════════

  // 1. sheets 不能为空
  if (!wb.sheets || wb.sheets.length === 0) {
    issues.push({ level: 'error', path: 'sheets', message: 'at least one sheet is required' });
  }

  const ssTexts = new Set((wb.sharedStrings || []).map(e => e.text));

  for (let i = 0; i < (wb.sheets || []).length; i++) {
    const sheet = wb.sheets[i];
    const path = `sheets[${i}]`;

    // 2. sheet name 不能为空
    if (!sheet.name) {
      issues.push({ level: 'error', path, message: 'sheet name cannot be empty' });
    }

    // 3. sharedString 引用的文本必须在 sharedStrings 中存在
    for (let r = 0; r < sheet.cells.length; r++) {
      for (let c = 0; c < sheet.cells[r].length; c++) {
        const cell = sheet.cells[r][c];
        if (!cell) continue;
        if (cell.type === 'sharedString' && typeof cell.value === 'string') {
          if (!ssTexts.has(cell.value)) {
            issues.push({
              level: 'error',
              path: `${path}.cells[${r}][${c}]`,
              message: `sharedString "${cell.value}" not found in sharedStrings`,
            });
          }
        }
      }
    }

    // 4. mergedCells 范围合法
    for (let j = 0; j < (sheet.mergedCells || []).length; j++) {
      const mc = sheet.mergedCells[j];
      if (mc.startRow > mc.endRow || mc.startCol > mc.endCol) {
        issues.push({
          level: 'error',
          path: `${path}.mergedCells[${j}]`,
          message: `invalid range: start(${mc.startRow},${mc.startCol}) > end(${mc.endRow},${mc.endCol})`,
        });
      }
      if (mc.startRow < 0 || mc.startCol < 0) {
        issues.push({
          level: 'error',
          path: `${path}.mergedCells[${j}]`,
          message: `negative start index: (${mc.startRow},${mc.startCol})`,
        });
      }
    }

    // 5. hyperlinks 必须有 ref 和 url
    for (let j = 0; j < (sheet.hyperlinks || []).length; j++) {
      const hl = sheet.hyperlinks[j];
      if (!hl.ref) {
        issues.push({ level: 'error', path: `${path}.hyperlinks[${j}]`, message: 'hyperlink missing ref' });
      }
      if (!hl.url) {
        issues.push({ level: 'error', path: `${path}.hyperlinks[${j}]`, message: 'hyperlink missing url' });
      }
    }

    // 6. table 必须有 ref 和 displayName
    for (let j = 0; j < (sheet.tables || []).length; j++) {
      const tbl = sheet.tables![j];
      if (!tbl.ref) {
        issues.push({ level: 'error', path: `${path}.tables[${j}]`, message: 'table missing ref' });
      }
      if (!tbl.displayName) {
        issues.push({ level: 'error', path: `${path}.tables[${j}]`, message: 'table missing displayName' });
      }
    }
  }

  // ═══════════════ 可选项 (warning) ═══════════════

  // 7. sharedStrings 建议非空（有 sharedString cell 时）
  let hasSharedStringCell = false;
  for (const sheet of wb.sheets || []) {
    for (const row of sheet.cells) {
      for (const cell of row) {
        if (cell?.type === 'sharedString') { hasSharedStringCell = true; break; }
      }
      if (hasSharedStringCell) break;
    }
    if (hasSharedStringCell) break;
  }
  if (hasSharedStringCell && (!wb.sharedStrings || wb.sharedStrings.length === 0)) {
    issues.push({ level: 'warning', path: 'sharedStrings', message: 'cells reference sharedString but sharedStrings array is empty' });
  }

  return issues;
}
