import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { NumberingDefinitions, NumberingLevel } from '../types.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function serializeNumbering(defs: NumberingDefinitions): string {
  const children: ParsedNode[] = [];

  for (const abstractNum of defs.abstractNums) {
    const lvlNodes: ParsedNode[] = abstractNum.levels.map(serializeLevel);
    children.push({
      tag: 'w:abstractNum',
      attrs: { 'w:abstractNumId': abstractNum.id },
      children: lvlNodes,
    });
  }

  for (const num of defs.nums) {
    children.push({
      tag: 'w:num',
      attrs: { 'w:numId': num.id },
      children: [
        { tag: 'w:abstractNumId', attrs: { 'w:val': num.abstractNumId }, children: [] },
      ],
    });
  }

  const root: ParsedNode = {
    tag: 'w:numbering',
    attrs: {
      'xmlns:w': W_NS,
    },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeLevel(lvl: NumberingLevel): ParsedNode {
  const children: ParsedNode[] = [
    { tag: 'w:start', attrs: { 'w:val': String(lvl.start) }, children: [] },
    { tag: 'w:numFmt', attrs: { 'w:val': lvl.numFmt }, children: [] },
    { tag: 'w:lvlText', attrs: { 'w:val': lvl.lvlText }, children: [] },
  ];

  if (lvl.alignment) {
    children.push({ tag: 'w:lvlJc', attrs: { 'w:val': lvl.alignment }, children: [] });
  }

  if (lvl.indent) {
    const attrs: Record<string, string> = {};
    if (lvl.indent.left !== undefined) attrs['w:left'] = String(lvl.indent.left);
    if (lvl.indent.hanging !== undefined) attrs['w:hanging'] = String(lvl.indent.hanging);
    children.push({
      tag: 'w:pPr',
      attrs: {},
      children: [{ tag: 'w:ind', attrs, children: [] }],
    });
  }

  return {
    tag: 'w:lvl',
    attrs: { 'w:ilvl': String(lvl.level) },
    children,
  };
}
