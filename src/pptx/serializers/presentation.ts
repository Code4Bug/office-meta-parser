import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { PptxPresentation, Slide } from '../types.js';
import { serializeSlide } from './slide.js';

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const NS = { P_NS, A_NS, R_NS };

export function serializePresentation(pres: PptxPresentation): string {
  const sldIdChildren = pres.slides.map((_slide, i) => ({
    tag: 'p:sldId',
    attrs: { id: String(256 + i), 'r:id': `rId${i + 3}` },
    children: [],
  }));

  const root: ParsedNode = {
    tag: 'p:presentation',
    attrs: { 'xmlns:p': P_NS, 'xmlns:a': A_NS, 'xmlns:r': R_NS },
    children: [
      {
        tag: 'p:sldMasterIdLst',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'p:sldMasterId',
            attrs: { id: '2147483648', 'r:id': 'rId2' },
            children: [],
          },
        ],
      },
      { tag: 'p:sldIdLst', attrs: {} as Record<string, string>, children: sldIdChildren },
      { tag: 'p:sldSz', attrs: { cx: String(pres.slideSize?.width ?? 9144000), cy: String(pres.slideSize?.height ?? 6858000) }, children: [] },
      { tag: 'p:notesSz', attrs: { cx: String(pres.notesSize?.width ?? 6858000), cy: String(pres.notesSize?.height ?? 9144000) }, children: [] },
    ],
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

export function serializePresentationRels(slides: Slide[]): string {
  const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

  const children: ParsedNode[] = [
    // Theme relationship
    {
      tag: 'Relationship',
      attrs: {
        Id: 'rId1',
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
        Target: 'theme/theme1.xml',
      },
      children: [],
    },
    // Slide master relationship
    {
      tag: 'Relationship',
      attrs: {
        Id: 'rId2',
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
        Target: 'slideMasters/slideMaster1.xml',
      },
      children: [],
    },
  ];

  // Slide relationships
  slides.forEach((_slide, i) => {
    children.push({
      tag: 'Relationship',
      attrs: {
        Id: `rId${i + 3}`,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
        Target: `slides/slide${i + 1}.xml`,
      },
      children: [],
    });
  });

  const root: ParsedNode = {
    tag: 'Relationships',
    attrs: { xmlns: RELS_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

export function serializeSlideRels(_index: number, hasComments?: boolean): string {
  const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

  const children: ParsedNode[] = [
    {
      tag: 'Relationship',
      attrs: {
        Id: 'rId1',
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
        Target: '../slideLayouts/slideLayout1.xml',
      },
      children: [],
    },
  ];

  if (hasComments) {
    children.push({
      tag: 'Relationship',
      attrs: {
        Id: 'rId2',
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
        Target: '../comments/comment' + (_index + 1) + '.xml',
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

export function serializeContentTypes(slides: Slide[], layoutCount = 1): string {
  const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

  const children: ParsedNode[] = [
    { tag: 'Default', attrs: { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }, children: [] },
    { tag: 'Default', attrs: { Extension: 'xml', ContentType: 'application/xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/ppt/presentation.xml', ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/ppt/presProps.xml', ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/ppt/theme/theme1.xml', ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml' }, children: [] },
    { tag: 'Override', attrs: { PartName: '/ppt/slideMasters/slideMaster1.xml', ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml' }, children: [] },
  ];

  for (let i = 1; i <= layoutCount; i++) {
    children.push({
      tag: 'Override',
      attrs: { PartName: `/ppt/slideLayouts/slideLayout${i}.xml`, ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml' },
      children: [],
    });
  }

  slides.forEach((_slide, i) => {
    children.push({
      tag: 'Override',
      attrs: { PartName: `/ppt/slides/slide${i + 1}.xml`, ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml' },
      children: [],
    });
  });

  const root: ParsedNode = {
    tag: 'Types',
    attrs: { xmlns: CT_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
