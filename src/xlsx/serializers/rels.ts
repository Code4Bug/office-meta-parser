import { serializeXml } from '../../core/xml.js';
import type { ParsedNode, ContentType } from '../../core/types.js';
import type { Sheet } from '../types.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

export function serializeWorkbookRels(sheets: Sheet[]): string {
  const children = sheets.map((_sheet, i) => ({
    tag: 'Relationship',
    attrs: {
      Id: `rId${i + 1}`,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
      Target: `worksheets/sheet${i + 1}.xml`,
    },
    children: [],
  }));

  // Add shared strings relationship
  children.push({
    tag: 'Relationship',
    attrs: {
      Id: `rId${sheets.length + 1}`,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
      Target: 'sharedStrings.xml',
    },
    children: [],
  });

  // Add styles relationship
  children.push({
    tag: 'Relationship',
    attrs: {
      Id: `rId${sheets.length + 2}`,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      Target: 'styles.xml',
    },
    children: [],
  });

  const root: ParsedNode = {
    tag: 'Relationships',
    attrs: { xmlns: RELS_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

export function serializeSheetRels(sheet: Sheet, index: number): string {
  const children: ParsedNode[] = [];
  let rid = 1;

  // Table relationships
  if (sheet.tables && sheet.tables.length > 0) {
    for (let i = 0; i < sheet.tables.length; i++) {
      children.push({
        tag: 'Relationship',
        attrs: {
          Id: `rId${rid++}`,
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
          Target: `../tables/table${sheet.tables[i].id}.xml`,
        },
        children: [],
      });
    }
  }

  // Comments relationship
  if (sheet.comments && sheet.comments.length > 0) {
    children.push({
      tag: 'Relationship',
      attrs: {
        Id: `rId${rid++}`,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
        Target: `../comments${index + 1}.xml`,
      },
      children: [],
    });
    children.push({
      tag: 'Relationship',
      attrs: {
        Id: `rId${rid++}`,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
        Target: `../drawings/vmlDrawing${index + 1}.vml`,
      },
      children: [],
    });
  }

  // Image relationship (existing)
  if (sheet.images && sheet.images.length > 0) {
    children.push({
      tag: 'Relationship',
      attrs: {
        Id: 'rId' + rid++,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
        Target: `../drawings/drawing${index + 1}.xml`,
      },
      children: [],
    });
  }

  const root: ParsedNode = {
    tag: 'Relationships',
    attrs: { xmlns: RELS_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

export function serializeContentTypes(sheets: Sheet[], extraContentTypes?: ContentType[]): string {
  const children: ParsedNode[] = [
    { tag: 'Default', attrs: { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }, children: [] },
    { tag: 'Default', attrs: { Extension: 'xml', ContentType: 'application/xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/xl/workbook.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/xl/sharedStrings.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/xl/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' }, children: [] },
  ];

  sheets.forEach((_sheet, i) => {
    children.push({
      tag: 'Override',
      attrs: { PartName: `/xl/worksheets/sheet${i + 1}.xml`, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' },
      children: [],
    });
  });

  // 合并 extra content types（以生成的为准，补充缺失的 Override）
  if (extraContentTypes) {
    const generatedPartNames = new Set<string>();
    for (const child of children) {
      if (child.attrs.PartName) generatedPartNames.add(child.attrs.PartName);
    }
    for (const ct of extraContentTypes) {
      if (ct.partName && !generatedPartNames.has(ct.partName)) {
        children.push({ tag: 'Override', attrs: { PartName: ct.partName, ContentType: ct.contentType }, children: [] });
        generatedPartNames.add(ct.partName);
      }
    }
  }

  const root: ParsedNode = {
    tag: 'Types',
    attrs: { xmlns: CT_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

export function serializeCommentsXml(sheetIndex: number, authorName: string, comments: import('../types.js').SheetComment[]): string {
  const authors = [authorName];
  const authorsNode = {
    tag: 'authors',
    attrs: {},
    children: [{ tag: 'author', attrs: {}, children: [authorName] }],
  };

  const commentList = comments.map(c => ({
    tag: 'comment',
    attrs: { ref: c.ref, authorId: String(c.authorId), shapeId: '0' },
    children: [{
      tag: 'text',
      attrs: {},
      children: [{ tag: 'r', attrs: {}, children: [{ tag: 't', attrs: {}, children: [c.text] }] }],
    }],
  }));

  const root: ParsedNode = {
    tag: 'comments',
    attrs: { xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' },
    children: [
      authorsNode,
      { tag: 'commentList', attrs: {}, children: commentList },
    ],
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
