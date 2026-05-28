import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { DocxDocument, Hyperlink, SectionProperties } from '../types.js';
import { serializeParagraph, serializeRun } from './paragraph.js';
import { serializeTable } from './table.js';
import { serializeImageParagraph } from './image.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const NS = { W_NS, R_NS };

export function semanticToXml(doc: DocxDocument): string {
  // 合并原始根元素属性（保留命名空间声明、mc:Ignorable 等）
  const attrs: Record<string, string> = {
    'xmlns:w': W_NS,
    'xmlns:r': R_NS,
  };
  if (doc.originalRootAttrs) {
    for (const [key, val] of Object.entries(doc.originalRootAttrs)) {
      if (!(key in attrs)) {
        attrs[key] = val;
      }
    }
  }

  const root: ParsedNode = {
    tag: 'w:document',
    attrs,
    children: [
      serializeBody(doc),
    ],
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeBody(doc: DocxDocument): ParsedNode {
  const children: ParsedNode[] = [];

  for (const block of doc.body.blocks) {
    if (block.type === 'paragraph') {
      children.push(serializeParagraph(block));
    } else if (block.type === 'table') {
      children.push(serializeTable(block));
    } else if (block.type === 'image') {
      children.push(serializeImageParagraph(block));
    } else if (block.type === 'hyperlink') {
      children.push(serializeHyperlink(block));
    } else if (block.type === 'bookmarkStart') {
      children.push({
        tag: 'w:bookmarkStart',
        attrs: { 'w:id': block.id, 'w:name': block.name },
        children: [],
      });
    } else if (block.type === 'bookmarkEnd') {
      children.push({
        tag: 'w:bookmarkEnd',
        attrs: { 'w:id': block.id },
        children: [],
      });
    }
  }

  // 添加节属性
  if (doc.body.sectionProperties) {
    children.push(serializeSectionProperties(doc.body.sectionProperties));
  }

  return {
    tag: 'w:body',
    attrs: {},
    children,
  };
}

function serializeSectionProperties(props: SectionProperties): ParsedNode {
  const children: ParsedNode[] = [];

  // 添加页眉引用
  if (props.headerReferenceId) {
    children.push({
      tag: 'w:headerReference',
      attrs: {
        'r:id': props.headerReferenceId,
        'w:type': props.headerReferenceType || 'default',
      },
      children: [],
    });
  }

  // 添加页脚引用
  if (props.footerReferenceId) {
    children.push({
      tag: 'w:footerReference',
      attrs: {
        'r:id': props.footerReferenceId,
        'w:type': props.footerReferenceType || 'default',
      },
      children: [],
    });
  }

  // 页面尺寸
  if (props.pageWidth || props.pageHeight || props.orientation) {
    const attrs: Record<string, string> = {};
    if (props.pageWidth) attrs['w:w'] = String(props.pageWidth);
    if (props.pageHeight) attrs['w:h'] = String(props.pageHeight);
    if (props.orientation) attrs['w:orient'] = props.orientation;
    children.push({ tag: 'w:pgSz', attrs, children: [] });
  }

  // 页面边距
  if (props.marginTop !== undefined || props.marginRight !== undefined ||
      props.marginBottom !== undefined || props.marginLeft !== undefined ||
      props.headerMargin !== undefined || props.footerMargin !== undefined ||
      props.gutter !== undefined) {
    const attrs: Record<string, string> = {};
    if (props.marginTop !== undefined) attrs['w:top'] = String(props.marginTop);
    if (props.marginRight !== undefined) attrs['w:right'] = String(props.marginRight);
    if (props.marginBottom !== undefined) attrs['w:bottom'] = String(props.marginBottom);
    if (props.marginLeft !== undefined) attrs['w:left'] = String(props.marginLeft);
    if (props.headerMargin !== undefined) attrs['w:header'] = String(props.headerMargin);
    if (props.footerMargin !== undefined) attrs['w:footer'] = String(props.footerMargin);
    if (props.gutter !== undefined) attrs['w:gutter'] = String(props.gutter);
    children.push({ tag: 'w:pgMar', attrs, children: [] });
  }

  // 分栏
  if (props.columnSpace || props.columnCount) {
    const attrs: Record<string, string> = {};
    if (props.columnSpace) attrs['w:space'] = String(props.columnSpace);
    if (props.columnCount) attrs['w:num'] = String(props.columnCount);
    children.push({ tag: 'w:cols', attrs, children: [] });
  }

  // 页码格式
  if (props.pageNumberFormat !== undefined || props.pageNumberStart !== undefined) {
    const attrs: Record<string, string> = {};
    if (props.pageNumberFormat) attrs['w:fmt'] = props.pageNumberFormat;
    if (props.pageNumberStart !== undefined) attrs['w:start'] = String(props.pageNumberStart);
    children.push({ tag: 'w:pgNumType', attrs, children: [] });
  }

  // 首页不同
  if (props.titlePage) {
    children.push({ tag: 'w:titlePg', attrs: {}, children: [] });
  }

  // 奇偶页不同
  if (props.evenAndOddHeaders) {
    children.push({ tag: 'w:evenAndOddHeaders', attrs: {}, children: [] });
  }

  // 垂直对齐
  if (props.verticalAlign) {
    children.push({ tag: 'w:vAlign', attrs: { 'w:val': props.verticalAlign }, children: [] });
  }

  return { tag: 'w:sectPr', attrs: {}, children };
}

function serializeHyperlink(hyperlink: Hyperlink): ParsedNode {
  const children: ParsedNode[] = [];

  for (const run of hyperlink.runs) {
    children.push(serializeRun(run));
  }

  const attrs: Record<string, string> = {
    'r:id': hyperlink.relationshipId,
  };

  if (hyperlink.tooltip) {
    attrs['w:tooltip'] = hyperlink.tooltip;
  }

  return {
    tag: 'w:hyperlink',
    attrs,
    children,
  };
}
