import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type {
  StyleDefinitions,
  ParagraphStyle,
  CharacterStyle,
  TableStyle,
  RunProperties,
  LatentStyle,
} from '../types.js';
import { serializeParagraphProperties } from './paragraph.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function serializeStyles(styles: StyleDefinitions, originalRootAttrs?: Record<string, string>): string {
  const children: ParsedNode[] = [];

  // 序列化潜在样式
  if (styles.latentStyles && styles.latentStyles.length > 0) {
    children.push(serializeLatentStyles(styles.latentStyles));
  }

  for (const style of styles.paragraphStyles) {
    children.push(serializeParagraphStyle(style));
  }
  for (const style of styles.characterStyles) {
    children.push(serializeCharacterStyle(style));
  }
  for (const style of styles.tableStyles) {
    children.push(serializeTableStyle(style));
  }

  // 合并原始根元素属性（保留命名空间声明、mc:Ignorable 等）
  const attrs: Record<string, string> = {
    'xmlns:w': W_NS,
    'xmlns:r': R_NS,
  };
  if (originalRootAttrs) {
    for (const [key, val] of Object.entries(originalRootAttrs)) {
      if (!(key in attrs)) {
        attrs[key] = val;
      }
    }
  }

  const root: ParsedNode = {
    tag: 'w:styles',
    attrs,
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeLatentStyles(latentStyles: LatentStyle[]): ParsedNode {
  const children: ParsedNode[] = latentStyles.map(ls => {
    const attrs: Record<string, string> = { 'w:name': ls.name };
    if (ls.uiPriority !== undefined) attrs['w:uiPriority'] = String(ls.uiPriority);
    if (ls.semiHidden) attrs['w:semiHidden'] = '1';
    if (ls.unhideWhenUsed) attrs['w:unhideWhenUsed'] = '1';
    if (ls.qFormat) attrs['w:qFormat'] = '1';
    return { tag: 'w:lsdException', attrs, children: [] };
  });

  return {
    tag: 'w:latentStyles',
    attrs: {
      'w:count': String(latentStyles.length),
    },
    children,
  };
}

function serializeParagraphStyle(style: ParagraphStyle): ParsedNode {
  const children: ParsedNode[] = [];

  if (style.name) {
    children.push({ tag: 'w:name', attrs: { 'w:val': style.name }, children: [] });
  }
  if (style.basedOn) {
    children.push({ tag: 'w:basedOn', attrs: { 'w:val': style.basedOn }, children: [] });
  }
  if (style.next) {
    children.push({ tag: 'w:next', attrs: { 'w:val': style.next }, children: [] });
  }
  if (style.uiPriority !== undefined) {
    children.push({ tag: 'w:uiPriority', attrs: { 'w:val': String(style.uiPriority) }, children: [] });
  }
  if (style.semiHidden) {
    children.push({ tag: 'w:semiHidden', attrs: {}, children: [] });
  }
  if (style.unhideWhenUsed) {
    children.push({ tag: 'w:unhideWhenUsed', attrs: {}, children: [] });
  }
  if (style.qFormat) {
    children.push({ tag: 'w:qFormat', attrs: {}, children: [] });
  }
  if (style.properties) {
    children.push(serializeParagraphProperties(style.properties));
  }
  if (style.runProperties) {
    children.push(serializeRunPropertiesFromStyle(style.runProperties));
  }

  const attrs: Record<string, string> = { 'w:type': 'paragraph', 'w:styleId': style.id };
  if (style.isDefault) attrs['w:default'] = '1';

  return {
    tag: 'w:style',
    attrs,
    children,
  };
}

function serializeCharacterStyle(style: CharacterStyle): ParsedNode {
  const children: ParsedNode[] = [];

  if (style.name) {
    children.push({ tag: 'w:name', attrs: { 'w:val': style.name }, children: [] });
  }
  if (style.basedOn) {
    children.push({ tag: 'w:basedOn', attrs: { 'w:val': style.basedOn }, children: [] });
  }
  if (style.properties) {
    children.push(serializeRunPropertiesFromStyle(style.properties));
  }

  return {
    tag: 'w:style',
    attrs: { 'w:type': 'character', 'w:styleId': style.id },
    children,
  };
}

function serializeTableStyle(style: TableStyle): ParsedNode {
  const children: ParsedNode[] = [];

  if (style.name) {
    children.push({ tag: 'w:name', attrs: { 'w:val': style.name }, children: [] });
  }
  if (style.basedOn) {
    children.push({ tag: 'w:basedOn', attrs: { 'w:val': style.basedOn }, children: [] });
  }

  return {
    tag: 'w:style',
    attrs: { 'w:type': 'table', 'w:styleId': style.id },
    children,
  };
}

function serializeRunPropertiesFromStyle(props: RunProperties): ParsedNode {
  const children: ParsedNode[] = [];

  if (props.fontFamily || props.fontFamilyEastAsia) {
    const rFontsAttrs: Record<string, string> = {};
    if (props.fontFamily) {
      rFontsAttrs['w:ascii'] = props.fontFamily;
      rFontsAttrs['w:hAnsi'] = props.fontFamily;
    }
    if (props.fontFamilyEastAsia) {
      rFontsAttrs['w:eastAsia'] = props.fontFamilyEastAsia;
    }
    children.push({ tag: 'w:rFonts', attrs: rFontsAttrs, children: [] });
  }
  if (props.bold) {
    children.push({ tag: 'w:b', attrs: {}, children: [] });
  }
  if (props.italic) {
    children.push({ tag: 'w:i', attrs: {}, children: [] });
  }
  if (props.underline) {
    children.push({ tag: 'w:u', attrs: { 'w:val': 'single' }, children: [] });
  }
  if (props.strike) {
    children.push({ tag: 'w:strike', attrs: {}, children: [] });
  }
  if (props.fontSize) {
    children.push({ tag: 'w:sz', attrs: { 'w:val': String(props.fontSize) }, children: [] });
  }
  if (props.fontSizeCs) {
    children.push({ tag: 'w:szCs', attrs: { 'w:val': String(props.fontSizeCs) }, children: [] });
  }
  if (props.color) {
    children.push({ tag: 'w:color', attrs: { 'w:val': props.color }, children: [] });
  }

  return { tag: 'w:rPr', attrs: {}, children };
}
