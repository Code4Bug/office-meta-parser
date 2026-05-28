import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { StyleDefinitions, ParagraphStyle, CharacterStyle, TableStyle, ParagraphProperties, RunProperties, LatentStyle } from '../types.js';
import { findChild } from './utils.js';

export function extractStyles(raw: RawDocument): StyleDefinitions {
  const stylesXml = raw.parts.get('word/styles.xml');
  if (!stylesXml) {
    return { paragraphStyles: [], characterStyles: [], tableStyles: [] };
  }

  const paragraphStyles: ParagraphStyle[] = [];
  const characterStyles: CharacterStyle[] = [];
  const tableStyles: TableStyle[] = [];
  let latentStyles: LatentStyle[] | undefined;

  for (const child of stylesXml.children) {
    if (typeof child === 'string') continue;

    if (child.tag === 'w:latentStyles') {
      latentStyles = parseLatentStyles(child);
      continue;
    }

    if (child.tag !== 'w:style') continue;

    const type = child.attrs['w:type'] || '';
    const styleId = child.attrs['w:styleId'] || '';

    if (type === 'paragraph') {
      paragraphStyles.push(parseParagraphStyle(child, styleId));
    } else if (type === 'character') {
      characterStyles.push(parseCharacterStyle(child, styleId));
    } else if (type === 'table') {
      tableStyles.push(parseTableStyle(child, styleId));
    }
  }

  const result: StyleDefinitions = { paragraphStyles, characterStyles, tableStyles };
  if (latentStyles) result.latentStyles = latentStyles;
  return result;
}

function parseLatentStyles(node: ParsedNode): LatentStyle[] {
  const result: LatentStyle[] = [];

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:lsdException') {
      const ls: LatentStyle = { name: child.attrs['w:name'] || '' };
      if (child.attrs['w:uiPriority']) ls.uiPriority = parseInt(child.attrs['w:uiPriority'], 10);
      if (child.attrs['w:semiHidden'] === '1' || child.attrs['w:semiHidden'] === 'true') ls.semiHidden = true;
      if (child.attrs['w:unhideWhenUsed'] === '1' || child.attrs['w:unhideWhenUsed'] === 'true') ls.unhideWhenUsed = true;
      if (child.attrs['w:qFormat'] === '1' || child.attrs['w:qFormat'] === 'true') ls.qFormat = true;
      result.push(ls);
    }
  }

  return result;
}

function parseParagraphStyle(node: ParsedNode, id: string): ParagraphStyle {
  const style: ParagraphStyle = { id };

  // 解析默认样式标记
  if (node.attrs['w:default'] === '1' || node.attrs['w:default'] === 'true') {
    style.isDefault = true;
  }

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:name': style.name = child.attrs['w:val']; break;
      case 'w:basedOn': style.basedOn = child.attrs['w:val']; break;
      case 'w:next': style.next = child.attrs['w:val']; break;
      case 'w:uiPriority':
        if (child.attrs['w:val']) style.uiPriority = parseInt(child.attrs['w:val'], 10);
        break;
      case 'w:semiHidden': style.semiHidden = true; break;
      case 'w:unhideWhenUsed': style.unhideWhenUsed = true; break;
      case 'w:qFormat': style.qFormat = true; break;
      case 'w:pPr': style.properties = parseParagraphProperties(child); break;
      case 'w:rPr': style.runProperties = parseRunProperties(child); break;
    }
  }

  return style;
}

function parseCharacterStyle(node: ParsedNode, id: string): CharacterStyle {
  const style: CharacterStyle = { id };

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:name': style.name = child.attrs['w:val']; break;
      case 'w:basedOn': style.basedOn = child.attrs['w:val']; break;
      case 'w:rPr': style.properties = parseRunProperties(child); break;
    }
  }

  return style;
}

function parseTableStyle(node: ParsedNode, id: string): TableStyle {
  const style: TableStyle = { id };

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:name') {
      style.name = child.attrs['w:val'];
    } else if (child.tag === 'w:basedOn') {
      style.basedOn = child.attrs['w:val'];
    }
  }

  return style;
}

function parseParagraphProperties(node: ParsedNode): ParagraphProperties {
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
        if (child.attrs['w:firstLine']) props.indent.firstLine = parseInt(child.attrs['w:firstLine'], 10);
        break;
      case 'w:spacing':
        props.spacing = {};
        if (child.attrs['w:before']) props.spacing.before = parseInt(child.attrs['w:before'], 10);
        if (child.attrs['w:after']) props.spacing.after = parseInt(child.attrs['w:after'], 10);
        if (child.attrs['w:line']) props.spacing.line = parseInt(child.attrs['w:line'], 10);
        break;
    }
  }

  return props;
}

function parseRunProperties(node: ParsedNode): RunProperties {
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
      case 'w:szCs':
        if (child.attrs['w:val']) props.fontSizeCs = parseInt(child.attrs['w:val'], 10);
        break;
      case 'w:color':
        if (child.attrs['w:val']) props.color = child.attrs['w:val'];
        break;
      case 'w:rFonts':
        if (child.attrs['w:ascii']) props.fontFamily = child.attrs['w:ascii'];
        if (child.attrs['w:eastAsia']) props.fontFamilyEastAsia = child.attrs['w:eastAsia'];
        break;
    }
  }

  return props;
}
