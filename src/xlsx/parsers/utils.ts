import type { ParsedNode } from '../../core/types.js';

export function findChild(node: ParsedNode, tag: string): ParsedNode | undefined {
  for (const child of node.children) {
    if (typeof child !== 'string' && child.tag === tag) {
      return child;
    }
  }
  return undefined;
}

export function getTextContent(node: ParsedNode): string {
  let result = '';
  for (const child of node.children) {
    if (typeof child === 'string') {
      result += child;
    } else {
      result += getTextContent(child);
    }
  }
  return result;
}

export function parseCellRef(ref: string): { row: number; col: number } | null {
  const match = ref.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;

  const colStr = match[1].toUpperCase();
  const rowStr = match[2];

  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col -= 1; // 0-based

  const row = parseInt(rowStr, 10) - 1; // 0-based
  if (isNaN(row)) return null;

  return { row, col };
}
