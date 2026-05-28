import type { ParsedNode } from '../../core/types.js';
import type { Paragraph, ParagraphProperties, TextRun, RunProperties, NumberingProperties, Field } from '../types.js';
import { findChild, getTextContent } from './utils.js';

export function parseParagraph(node: ParsedNode): Paragraph {
  const runs: TextRun[] = [];
  let properties: ParagraphProperties | undefined;
  let style: string | undefined;
  let numbering: NumberingProperties | undefined;

  // 用于解析字段
  let inField = false;
  let fieldInstruction = '';
  let fieldResult = '';

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:r') {
      const run = parseRunWithField(child, inField);
      if (run.fieldStart) {
        inField = true;
        fieldInstruction = '';
        fieldResult = '';
      } else if (run.fieldSeparate) {
        // separate 标记，不做处理
      } else if (run.fieldEnd) {
        inField = false;
        const fieldRun: TextRun = {
          text: fieldResult || '',
          field: {
            instruction: fieldInstruction.trim(),
            result: fieldResult,
          },
          ...run.properties,
        };
        runs.push(fieldRun);
      } else if (run.instruction) {
        fieldInstruction += run.instruction;
      } else if (inField && run.textRun?.text) {
        fieldResult += run.textRun.text;
      } else if (run.textRun) {
        runs.push(run.textRun);
      }
    } else if (child.tag === 'w:ins' || child.tag === 'w:del') {
      // 修订标记: w:ins (插入) / w:del (删除)
      const revType = child.tag === 'w:ins' ? 'insert' as const : 'delete' as const;
      const revAuthor = child.attrs['w:author'] || '';
      const revDate = child.attrs['w:date'] || '';
      for (const revChild of child.children) {
        if (typeof revChild === 'string') continue;
        if (revChild.tag === 'w:r') {
          const run = parseRunWithField(revChild, inField);
          if (run.textRun) {
            run.textRun.revisionType = revType;
            run.textRun.revisionAuthor = revAuthor;
            run.textRun.revisionDate = revDate;
            runs.push(run.textRun);
          }
        }
      }
    } else if (child.tag === 'w:pPr') {
      properties = parseParagraphProperties(child);
      const pStyle = findChild(child, 'w:pStyle');
      if (pStyle?.attrs['w:val']) {
        style = pStyle.attrs['w:val'];
      }
      const numPr = findChild(child, 'w:numPr');
      if (numPr) {
        numbering = parseNumberingProperties(numPr);
      }
    }
  }

  const result: Paragraph = { type: 'paragraph', runs };
  if (properties) result.properties = properties;
  if (style) result.style = style;
  if (numbering) result.numbering = numbering;
  return result;
}

export function parseRun(node: ParsedNode): TextRun {
  let text = '';
  let properties: RunProperties | undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:t' && text === '') {
      text = getTextContent(child);
    } else if (child.tag === 'w:rPr') {
      properties = parseRunProperties(child);
    }
  }

  return { text, ...properties };
}

interface FieldRunResult {
  fieldStart?: boolean;
  fieldEnd?: boolean;
  fieldSeparate?: boolean;
  instruction?: string;
  text?: string;
  textRun?: TextRun;
  properties?: RunProperties;
}

function parseRunWithField(node: ParsedNode, inField: boolean): FieldRunResult {
  let text = '';
  let properties: RunProperties | undefined;
  let fldCharType: string | undefined;
  let instruction: string | undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:t' && text === '') {
      text = getTextContent(child);
    } else if (child.tag === 'w:rPr') {
      properties = parseRunProperties(child);
    } else if (child.tag === 'w:fldChar') {
      fldCharType = child.attrs['w:fldCharType'];
    } else if (child.tag === 'w:instrText') {
      instruction = getTextContent(child);
    }
  }

  if (fldCharType === 'begin') {
    return { fieldStart: true, properties };
  } else if (fldCharType === 'separate') {
    return { fieldSeparate: true, properties };
  } else if (fldCharType === 'end') {
    return { fieldEnd: true, properties };
  } else if (instruction) {
    return { instruction, properties };
  }

  return {
    textRun: { text, ...properties },
    properties,
  };
}

export function parseParagraphProperties(node: ParsedNode): ParagraphProperties {
  const props: ParagraphProperties = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:jc':
        if (child.attrs['w:val']) {
          const val = child.attrs['w:val'];
          if (val === 'left' || val === 'center' || val === 'right' || val === 'both') {
            props.alignment = val === 'both' ? 'justify' : val;
          }
        }
        break;
      case 'w:ind':
        props.indent = {};
        if (child.attrs['w:left']) props.indent.left = parseInt(child.attrs['w:left'], 10);
        if (child.attrs['w:right']) props.indent.right = parseInt(child.attrs['w:right'], 10);
        if (child.attrs['w:start']) props.indent.start = parseInt(child.attrs['w:start'], 10);
        if (child.attrs['w:end']) props.indent.end = parseInt(child.attrs['w:end'], 10);
        if (child.attrs['w:firstLine']) props.indent.firstLine = parseInt(child.attrs['w:firstLine'], 10);
        if (child.attrs['w:hanging']) props.indent.hanging = parseInt(child.attrs['w:hanging'], 10);
        if (child.attrs['w:leftChars']) props.indent.leftChars = parseInt(child.attrs['w:leftChars'], 10);
        if (child.attrs['w:rightChars']) props.indent.rightChars = parseInt(child.attrs['w:rightChars'], 10);
        if (child.attrs['w:startChars']) props.indent.startChars = parseInt(child.attrs['w:startChars'], 10);
        if (child.attrs['w:endChars']) props.indent.endChars = parseInt(child.attrs['w:endChars'], 10);
        if (child.attrs['w:firstLineChars']) props.indent.firstLineChars = parseInt(child.attrs['w:firstLineChars'], 10);
        if (child.attrs['w:hangingChars']) props.indent.hangingChars = parseInt(child.attrs['w:hangingChars'], 10);
        break;
      case 'w:spacing':
        props.spacing = {};
        if (child.attrs['w:before']) props.spacing.before = parseInt(child.attrs['w:before'], 10);
        if (child.attrs['w:after']) props.spacing.after = parseInt(child.attrs['w:after'], 10);
        if (child.attrs['w:line']) props.spacing.line = parseInt(child.attrs['w:line'], 10);
        if (child.attrs['w:lineRule']) {
          const rule = child.attrs['w:lineRule'];
          if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') {
            props.spacing.lineRule = rule;
          }
        }
        break;
      case 'w:keepNext':
        props.keepNext = true;
        break;
      case 'w:keepLines':
        props.keepLines = true;
        break;
      case 'w:pageBreakBefore':
        props.pageBreakBefore = true;
        break;
      case 'w:outlineLvl':
        if (child.attrs['w:val'] !== undefined) {
          props.outlineLevel = parseInt(child.attrs['w:val'], 10);
        }
        break;
      case 'w:shd':
        props.shading = {};
        if (child.attrs['w:fill']) props.shading.fill = child.attrs['w:fill'];
        if (child.attrs['w:color']) props.shading.color = child.attrs['w:color'];
        if (child.attrs['w:val']) props.shading.pattern = child.attrs['w:val'];
        break;
      case 'w:pBdr': {
        props.border = {};
        for (const bdrChild of child.children) {
          if (typeof bdrChild === 'string') continue;
          const side = bdrChild.tag.replace('w:', '') as 'top' | 'bottom' | 'left' | 'right' | 'between';
          if (['top', 'bottom', 'left', 'right', 'between'].includes(side)) {
            props.border[side] = {
              style: bdrChild.attrs['w:val'] || 'single',
              ...(bdrChild.attrs['w:sz'] ? { size: parseInt(bdrChild.attrs['w:sz'], 10) } : {}),
              ...(bdrChild.attrs['w:color'] ? { color: bdrChild.attrs['w:color'] } : {}),
            };
          }
        }
        break;
      }
      case 'w:tabs': {
        props.tabs = [];
        for (const tabChild of child.children) {
          if (typeof tabChild === 'string') continue;
          if (tabChild.tag === 'w:tab' && tabChild.attrs['w:val'] && tabChild.attrs['w:pos']) {
            props.tabs.push({
              position: parseInt(tabChild.attrs['w:pos'], 10) || 0,
              alignment: tabChild.attrs['w:val'] as any,
              ...(tabChild.attrs['w:leader'] ? { leader: tabChild.attrs['w:leader'] as any } : {}),
            });
          }
        }
        break;
      }
      case 'w:pPrChange':
        props.formatChangeAuthor = child.attrs['w:author'] || '';
        props.formatChangeDate = child.attrs['w:date'] || '';
        break;
    }
  }

  return props;
}

function parseNumberingProperties(node: ParsedNode): NumberingProperties | undefined {
  const ilvl = findChild(node, 'w:ilvl');
  const numId = findChild(node, 'w:numId');

  if (!numId?.attrs['w:val']) return undefined;

  const level = ilvl?.attrs['w:val'] ? parseInt(ilvl.attrs['w:val'], 10) : 0;

  return {
    level,
    numId: numId.attrs['w:val'],
  };
}

export function parseRunProperties(node: ParsedNode): RunProperties {
  const props: RunProperties = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:b': props.bold = true; break;
      case 'w:i': props.italic = true; break;
      case 'w:u': props.underline = true; break;
      case 'w:strike': props.strike = true; break;
      case 'w:sz':
        if (child.attrs['w:val']) props.fontSize = parseInt(child.attrs['w:val'], 10);
        break;
      case 'w:color':
        if (child.attrs['w:val']) props.color = child.attrs['w:val'];
        break;
      case 'w:vertAlign':
        if (child.attrs['w:val'] === 'superscript') props.superscript = true;
        else if (child.attrs['w:val'] === 'subscript') props.subscript = true;
        break;
      case 'w:rFonts':
        if (child.attrs['w:ascii']) props.fontFamily = child.attrs['w:ascii'];
        break;
      case 'w:highlight':
        if (child.attrs['w:val']) props.highlight = child.attrs['w:val'];
        break;
      case 'w:shd':
        if (child.attrs['w:fill']) props.shadingColor = child.attrs['w:fill'];
        if (child.attrs['w:val']) props.shadingPattern = child.attrs['w:val'];
        break;
      case 'w:caps':
        props.caps = true;
        break;
      case 'w:smallCaps':
        props.smallCaps = true;
        break;
      case 'w:dstrike':
        props.dstrike = true;
        break;
      case 'w:vanish':
        props.vanish = true;
        break;
      case 'w:spacing':
        if (child.attrs['w:val']) props.characterSpacing = parseInt(child.attrs['w:val'], 10);
        break;
      case 'w:kern':
        if (child.attrs['w:val']) props.kern = parseInt(child.attrs['w:val'], 10);
        break;
      case 'w:em':
        if (child.attrs['w:val']) props.emphasis = child.attrs['w:val'];
        break;
      case 'w:bdr':
        props.characterBorder = {
          style: child.attrs['w:val'] || 'single',
          ...(child.attrs['w:sz'] ? { size: parseInt(child.attrs['w:sz'], 10) } : {}),
          ...(child.attrs['w:color'] ? { color: child.attrs['w:color'] } : {}),
        };
        break;
      case 'w:position':
        if (child.attrs['w:val']) props.verticalPosition = parseInt(child.attrs['w:val'], 10);
        break;
      case 'w:rPrChange':
        props.formatChangeAuthor = child.attrs['w:author'] || '';
        props.formatChangeDate = child.attrs['w:date'] || '';
        break;
    }
  }

  return props;
}
