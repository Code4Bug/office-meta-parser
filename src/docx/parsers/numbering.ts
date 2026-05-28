import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { NumberingDefinitions, AbstractNum, NumberingLevel, Num, RunProperties } from '../types.js';
import { findChild } from './utils.js';
import { parseRunProperties } from './paragraph.js';

export function extractNumbering(raw: RawDocument): NumberingDefinitions {
  const numberingXml = raw.parts.get('word/numbering.xml');
  if (!numberingXml) return { abstractNums: [], nums: [] };

  const abstractNums: AbstractNum[] = [];
  const nums: Num[] = [];

  for (const child of numberingXml.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:abstractNum') {
      abstractNums.push(parseAbstractNum(child));
    } else if (child.tag === 'w:num') {
      const id = child.attrs['w:numId'];
      const abstractId = findChild(child, 'w:abstractNumId');
      if (id && abstractId?.attrs['w:val']) {
        nums.push({ id, abstractNumId: abstractId.attrs['w:val'] });
      }
    }
  }

  return { abstractNums, nums };
}

function parseAbstractNum(node: ParsedNode): AbstractNum {
  const id = node.attrs['w:abstractNumId'] || '';
  const levels: NumberingLevel[] = [];

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:lvl') {
      levels.push(parseLevel(child));
    }
  }

  return { id, levels };
}

function parseLevel(node: ParsedNode): NumberingLevel {
  const level = parseInt(node.attrs['w:ilvl'] || '0', 10);
  let numFmt = 'decimal';
  let lvlText = '';
  let start = 1;
  let indent: NumberingLevel['indent'];
  let alignment: string | undefined;
  let runProperties: RunProperties | undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:numFmt':
        if (child.attrs['w:val']) numFmt = child.attrs['w:val'];
        break;
      case 'w:lvlText':
        if (child.attrs['w:val'] !== undefined) lvlText = child.attrs['w:val'];
        break;
      case 'w:start':
        if (child.attrs['w:val']) start = parseInt(child.attrs['w:val'], 10);
        break;
      case 'w:lvlJc':
        if (child.attrs['w:val']) alignment = child.attrs['w:val'];
        break;
      case 'w:pPr': {
        const ind = findChild(child, 'w:ind');
        if (ind) {
          indent = {};
          if (ind.attrs['w:left']) indent.left = parseInt(ind.attrs['w:left'], 10);
          if (ind.attrs['w:hanging']) indent.hanging = parseInt(ind.attrs['w:hanging'], 10);
        }
        break;
      }
      case 'w:rPr':
        runProperties = parseRunProperties(child);
        break;
    }
  }

  const result: NumberingLevel = { level, numFmt, lvlText, start };
  if (indent) result.indent = indent;
  if (alignment) result.alignment = alignment;
  if (runProperties) result.runProperties = runProperties;
  return result;
}
