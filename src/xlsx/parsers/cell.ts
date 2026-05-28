import type { ParsedNode } from '../../core/types.js';
import type { Cell, SharedStringEntry } from '../types.js';
import { findChild, getTextContent } from './utils.js';

// Common Excel built-in date format IDs
const DATE_FORMAT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22,
  27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47,
  56, 57, 58,
  64, 65, 66, 67,
]);

function isDateFormat(styleId: number): boolean {
  return DATE_FORMAT_IDS.has(styleId);
}

function excelDateToISOString(excelDate: number): string {
  const utcDays = Math.floor(excelDate - 25569);
  const utcValue = utcDays * 86400;
  const fractionalDay = excelDate - Math.floor(excelDate);

  let totalSeconds = Math.round(86400 * fractionalDay);
  const seconds = totalSeconds % 60;
  totalSeconds -= seconds;
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds - hours * 3600) / 60);

  const date = new Date(utcValue * 1000);
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');

  if (hours === 0 && minutes === 0 && seconds === 0) {
    return `${year}-${month}-${day}`;
  }

  return `${year}-${month}-${day}T${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

export function parseCell(node: ParsedNode, sharedStrings: SharedStringEntry[]): Cell {
  const type = node.attrs['t'] || '';
  const style = node.attrs['s'] || '';
  const vNode = findChild(node, 'v');
  const isNode = findChild(node, 'is');

  let value: string | number | boolean | null = null;
  let cellType: Cell['type'] = 'string';

  if (type === 's') {
    const idx = vNode ? parseInt(getTextContent(vNode), 10) : 0;
    value = sharedStrings[idx]?.text || '';
    cellType = 'sharedString';
  } else if (type === 'inlineStr') {
    if (isNode) {
      const tNode = findChild(isNode, 't');
      value = tNode ? getTextContent(tNode) : '';
    }
    cellType = 'string';
  } else if (type === 'b') {
    value = vNode ? getTextContent(vNode) === '1' : false;
    cellType = 'boolean';
  } else if (type === 'e') {
    value = vNode ? getTextContent(vNode) : null;
    cellType = 'error';
  } else if (vNode) {
    const rawValue = getTextContent(vNode);
    const num = Number(rawValue);
    if (!isNaN(num) && rawValue !== '') {
      value = num;
      cellType = 'number';
    } else {
      value = rawValue;
      cellType = 'string';
    }
  }

  if (cellType === 'number' && style && isDateFormat(parseInt(style, 10))) {
    cellType = 'date';
    value = excelDateToISOString(value as number);
  }

  const fNode = findChild(node, 'f');
  const formula = fNode ? getTextContent(fNode) : undefined;

  const cell: Cell = {
    value,
    formula,
    type: formula ? 'formula' : cellType,
  };

  // Formula type attributes
  if (fNode) {
    const fType = fNode.attrs['t'];
    if (fType === 'array') {
      cell.formulaType = 'array';
      if (fNode.attrs['ref']) cell.formulaRef = fNode.attrs['ref'];
    } else if (fType === 'shared') {
      cell.formulaType = 'shared';
      if (fNode.attrs['si']) cell.sharedFormulaIndex = parseInt(fNode.attrs['si'], 10);
    }
  }

  // Preserve style
  if (style) cell.style = style;

  return cell;
}
