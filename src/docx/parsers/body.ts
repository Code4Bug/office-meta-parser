import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { DocxBody, DocxBlock, SectionProperties } from '../types.js';
import { parseParagraph } from './paragraph.js';
import { parseTable } from './table.js';
import { parseHyperlink } from './hyperlink.js';
import { parseImage } from './image.js';
import { findChild } from './utils.js';

export function extractBody(raw: RawDocument): DocxBody {
  const documentXml = raw.parts.get('word/document.xml');
  if (!documentXml) return { blocks: [] };

  const body = findChild(documentXml, 'w:body');
  if (!body) return { blocks: [] };

  const blocks: DocxBlock[] = [];
  let sectionProperties: SectionProperties | undefined;

  for (const child of body.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:p') {
      const image = extractImageFromParagraph(raw, child);
      if (image) {
        blocks.push(image);
      } else {
        blocks.push(parseParagraph(child));
      }
    } else if (child.tag === 'w:tbl') {
      blocks.push(parseTable(child));
    } else if (child.tag === 'w:hyperlink') {
      blocks.push(parseHyperlink(raw, child));
    } else if (child.tag === 'w:bookmarkStart') {
      if (child.attrs['w:id'] && child.attrs['w:name']) {
        blocks.push({
          type: 'bookmarkStart',
          id: child.attrs['w:id'],
          name: child.attrs['w:name'],
        });
      }
    } else if (child.tag === 'w:bookmarkEnd') {
      if (child.attrs['w:id']) {
        blocks.push({ type: 'bookmarkEnd', id: child.attrs['w:id'] });
      }
    } else if (child.tag === 'w:sectPr') {
      sectionProperties = parseSectionProperties(child);
    }
  }

  const result: DocxBody = { blocks };
  if (sectionProperties) result.sectionProperties = sectionProperties;
  return result;
}

function parseSectionProperties(node: ParsedNode): SectionProperties {
  const props: SectionProperties = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:headerReference':
        if (child.attrs['r:id']) props.headerReferenceId = child.attrs['r:id'];
        if (child.attrs['w:type']) props.headerReferenceType = child.attrs['w:type'];
        break;
      case 'w:footerReference':
        if (child.attrs['r:id']) props.footerReferenceId = child.attrs['r:id'];
        if (child.attrs['w:type']) props.footerReferenceType = child.attrs['w:type'];
        break;
      case 'w:pgSz':
        if (child.attrs['w:w']) props.pageWidth = parseInt(child.attrs['w:w'], 10);
        if (child.attrs['w:h']) props.pageHeight = parseInt(child.attrs['w:h'], 10);
        if (child.attrs['w:orient']) {
          const o = child.attrs['w:orient'];
          if (o === 'portrait' || o === 'landscape') props.orientation = o;
        }
        break;
      case 'w:pgMar':
        if (child.attrs['w:top']) props.marginTop = parseInt(child.attrs['w:top'], 10);
        if (child.attrs['w:right']) props.marginRight = parseInt(child.attrs['w:right'], 10);
        if (child.attrs['w:bottom']) props.marginBottom = parseInt(child.attrs['w:bottom'], 10);
        if (child.attrs['w:left']) props.marginLeft = parseInt(child.attrs['w:left'], 10);
        if (child.attrs['w:header']) props.headerMargin = parseInt(child.attrs['w:header'], 10);
        if (child.attrs['w:footer']) props.footerMargin = parseInt(child.attrs['w:footer'], 10);
        if (child.attrs['w:gutter']) props.gutter = parseInt(child.attrs['w:gutter'], 10);
        break;
      case 'w:cols':
        if (child.attrs['w:space']) props.columnSpace = parseInt(child.attrs['w:space'], 10);
        if (child.attrs['w:num']) props.columnCount = parseInt(child.attrs['w:num'], 10);
        break;
      case 'w:pgNumType':
        if (child.attrs['w:fmt']) props.pageNumberFormat = child.attrs['w:fmt'];
        if (child.attrs['w:start']) props.pageNumberStart = parseInt(child.attrs['w:start'], 10);
        break;
      case 'w:titlePg':
        props.titlePage = true;
        break;
      case 'w:evenAndOddHeaders':
        props.evenAndOddHeaders = true;
        break;
      case 'w:vAlign':
        if (child.attrs['w:val']) {
          const v = child.attrs['w:val'];
          if (v === 'top' || v === 'center' || v === 'bottom' || v === 'both') {
            props.verticalAlign = v;
          }
        }
        break;
    }
  }

  return props;
}

function extractImageFromParagraph(raw: RawDocument, para: ParsedNode) {
  for (const child of para.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'w:r') continue;
    for (const runChild of child.children) {
      if (typeof runChild === 'string') continue;
      if (runChild.tag === 'w:drawing') {
        return parseImage(raw, runChild);
      }
    }
  }
  return null;
}
