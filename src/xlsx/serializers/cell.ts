import type { ParsedNode } from '../../core/types.js';
import type { Cell, SharedStringEntry } from '../types.js';

export function serializeCell(cell: Cell, rowIdx: number, colIdx: number, sharedStrings: SharedStringEntry[]): ParsedNode {
  const colLetter = getColLetter(colIdx);
  const cellRef = `${colLetter}${rowIdx + 1}`;

  const attrs: Record<string, string> = { r: cellRef };
  const children: ParsedNode[] = [];

  if (cell.type === 'sharedString' && typeof cell.value === 'string') {
    attrs.t = 's';
    const idx = sharedStrings.findIndex(e => e.text === cell.value);
    children.push({
      tag: 'v',
      attrs: {},
      children: [String(idx >= 0 ? idx : 0)],
    });
  } else if (cell.type === 'string' && typeof cell.value === 'string') {
    attrs.t = 'inlineStr';
    children.push({
      tag: 'is',
      attrs: {},
      children: [{ tag: 't', attrs: {}, children: [cell.value] }],
    });
  } else if (cell.type === 'number' && typeof cell.value === 'number') {
    children.push({
      tag: 'v',
      attrs: {},
      children: [String(cell.value)],
    });
  } else if (cell.type === 'boolean' && typeof cell.value === 'boolean') {
    attrs.t = 'b';
    children.push({
      tag: 'v',
      attrs: {},
      children: [cell.value ? '1' : '0'],
    });
  } else if (cell.value !== null && cell.value !== undefined) {
    children.push({
      tag: 'v',
      attrs: {},
      children: [String(cell.value)],
    });
  }

  if (cell.formula) {
    children.unshift({
      tag: 'f',
      attrs: {},
      children: [cell.formula],
    });
  }

  return { tag: 'c', attrs, children };
}

export function getColLetter(col: number): string {
  let result = '';
  let n = col;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}
