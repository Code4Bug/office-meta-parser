import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { SlideMaster, SlideLayout, Placeholder, MasterTxStyles, TextStyle, TextLevelStyle } from '../types.js';

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function serializeSlideMaster(master: SlideMaster, index: number): string {
  const xfrm: ParsedNode = {
    tag: 'a:xfrm',
    attrs: {} as Record<string, string>,
    children: [
      { tag: 'a:off', attrs: { x: '0', y: '0' }, children: [] },
      { tag: 'a:ext', attrs: { cx: '0', cy: '0' }, children: [] },
      { tag: 'a:chOff', attrs: { x: '0', y: '0' }, children: [] },
      { tag: 'a:chExt', attrs: { cx: '0', cy: '0' }, children: [] },
    ],
  };

  const bg: ParsedNode = master.background
    ? serializeBackground(master.background)
    : {
        tag: 'p:bg',
        attrs: {} as Record<string, string>,
        children: [{
          tag: 'p:bgRef',
          attrs: { idx: '1001' },
          children: [{ tag: 'a:schemeClr', attrs: { val: 'bg1' }, children: [] }],
        }],
      };

  const children: ParsedNode[] = [
    {
      tag: 'p:cSld',
      attrs: {} as Record<string, string>,
      children: [
        bg,
        {
          tag: 'p:spTree',
          attrs: {} as Record<string, string>,
          children: [
            {
              tag: 'p:nvGrpSpPr',
              attrs: {} as Record<string, string>,
              children: [
                { tag: 'p:cNvPr', attrs: { id: '1', name: '' }, children: [] },
                { tag: 'p:cNvGrpSpPr', attrs: {} as Record<string, string>, children: [] },
                { tag: 'p:nvPr', attrs: {} as Record<string, string>, children: [] },
              ],
            },
            { tag: 'p:grpSpPr', attrs: {} as Record<string, string>, children: [xfrm] },
          ],
        },
      ],
    },
    {
      tag: 'p:clrMap',
      attrs: {
        bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
        accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
        accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
        hlink: 'hlink', folHlink: 'folHlink',
      },
      children: [],
    },
    {
      tag: 'p:sldLayoutIdLst',
      attrs: {} as Record<string, string>,
      children: master.layouts.map((layout, i) => ({
        tag: 'p:sldLayoutId',
        attrs: { id: String(2147483649 + i), 'r:id': `rId${i + 1}` },
        children: [],
      })),
    },
    serializeTxStyles(master.txStyles),
  ];

  const root: ParsedNode = {
    tag: 'p:sldMaster',
    attrs: { 'xmlns:p': P_NS, 'xmlns:a': A_NS, 'xmlns:r': R_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

export function serializeSlideLayout(layout: SlideLayout): string {
  // 根 spTree 占用 id=1，子 shape 从 2 开始
  const spChildren: ParsedNode[] = layout.placeholders.map((ph, i) =>
    serializePlaceholderShape(ph, i + 2)
  );

  const xfrm: ParsedNode = {
    tag: 'a:xfrm',
    attrs: {} as Record<string, string>,
    children: [
      { tag: 'a:off', attrs: { x: '0', y: '0' }, children: [] },
      { tag: 'a:ext', attrs: { cx: '0', cy: '0' }, children: [] },
      { tag: 'a:chOff', attrs: { x: '0', y: '0' }, children: [] },
      { tag: 'a:chExt', attrs: { cx: '0', cy: '0' }, children: [] },
    ],
  };

  const attrs: Record<string, string> = {
    'xmlns:p': P_NS, 'xmlns:a': A_NS, 'xmlns:r': R_NS,
  };
  if (layout.type) attrs.type = layout.type;
  attrs.preserve = '1';

  const children: ParsedNode[] = [
    {
      tag: 'p:cSld',
      attrs: { name: layout.name || '' },
      children: [
        {
          tag: 'p:spTree',
          attrs: {} as Record<string, string>,
          children: [
            {
              tag: 'p:nvGrpSpPr',
              attrs: {} as Record<string, string>,
              children: [
                { tag: 'p:cNvPr', attrs: { id: '1', name: '' }, children: [] },
                { tag: 'p:cNvGrpSpPr', attrs: {} as Record<string, string>, children: [] },
                { tag: 'p:nvPr', attrs: {} as Record<string, string>, children: [] },
              ],
            },
            { tag: 'p:grpSpPr', attrs: {} as Record<string, string>, children: [xfrm] },
            ...spChildren,
          ],
        },
      ],
    },
    {
      tag: 'p:clrMapOvr',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:masterClrMapping', attrs: {} as Record<string, string>, children: [] }],
    },
  ];

  const root: ParsedNode = {
    tag: 'p:sldLayout',
    attrs,
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializePlaceholderShape(ph: Placeholder, id: number): ParsedNode {
  const phAttrs: Record<string, string> = { type: ph.type };
  if (ph.index !== undefined) {
    phAttrs.idx = String(ph.index);
  }

  return {
    tag: 'p:sp',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:nvSpPr',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'p:cNvPr', attrs: { id: String(id), name: `Placeholder ${id}` }, children: [] },
          { tag: 'p:cNvSpPr', attrs: {} as Record<string, string>, children: [] },
          {
            tag: 'p:nvPr',
            attrs: {} as Record<string, string>,
            children: [
              { tag: 'p:ph', attrs: phAttrs, children: [] },
            ],
          },
        ],
      },
      {
        tag: 'p:spPr',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'a:xfrm',
            attrs: {} as Record<string, string>,
            children: [
              { tag: 'a:off', attrs: { x: String(ph.position.x), y: String(ph.position.y) }, children: [] },
              { tag: 'a:ext', attrs: { cx: String(ph.position.width), cy: String(ph.position.height) }, children: [] },
            ],
          },
        ],
      },
      {
        tag: 'p:txBody',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'a:bodyPr', attrs: {} as Record<string, string>, children: [] },
          { tag: 'a:lstStyle', attrs: {} as Record<string, string>, children: [] },
          {
            tag: 'a:p',
            attrs: {} as Record<string, string>,
            children: [],
          },
        ],
      },
    ],
  };
}

function serializeTxStyles(txStyles?: MasterTxStyles): ParsedNode {
  const children: ParsedNode[] = [];

  if (txStyles?.titleStyle) {
    children.push(serializeTextStyle('p:titleStyle', txStyles.titleStyle));
  }
  if (txStyles?.bodyStyle) {
    children.push(serializeTextStyle('p:bodyStyle', txStyles.bodyStyle));
  }
  if (txStyles?.otherStyle) {
    children.push(serializeTextStyle('p:otherStyle', txStyles.otherStyle));
  }

  // Fallback: ensure 3 children if no txStyles data
  while (children.length < 3) {
    children.push({ tag: 'a:lstStyle', attrs: {} as Record<string, string>, children: [] });
  }

  return {
    tag: 'p:txStyles',
    attrs: {} as Record<string, string>,
    children,
  };
}

function serializeTextStyle(tag: string, style: TextStyle): ParsedNode {
  const children: ParsedNode[] = [];
  if (style.levels) {
    for (const level of style.levels) {
      children.push(serializeTextLevelStyle(level));
    }
  }
  return { tag, attrs: {} as Record<string, string>, children };
}

function serializeTextLevelStyle(level: TextLevelStyle): ParsedNode {
  const tag = `a:lvl${level.level}pPr`;
  const attrs: Record<string, string> = {};
  if (level.alignment) attrs.algn = level.alignment;
  if (level.marL !== undefined) attrs.marL = String(Math.round(level.marL * 12700));
  if (level.indent !== undefined) attrs.indent = String(Math.round(level.indent * 12700));

  const children: ParsedNode[] = [];
  if (level.spcBef !== undefined) {
    children.push({
      tag: 'a:spcBef',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:spcPts', attrs: { val: String(Math.round(level.spcBef * 100)) }, children: [] }],
    });
  }
  if (level.spcAft !== undefined) {
    children.push({
      tag: 'a:spcAft',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:spcPts', attrs: { val: String(Math.round(level.spcAft * 100)) }, children: [] }],
    });
  }
  if (level.lnSpc !== undefined) {
    children.push({
      tag: 'a:lnSpc',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:spcPct', attrs: { val: String(Math.round(level.lnSpc * 1000)) }, children: [] }],
    });
  }
  if (level.fontScale !== undefined) {
    children.push({
      tag: 'a:defRPr',
      attrs: { sz: String(Math.round(level.fontScale * 100)) },
      children: [],
    });
  }

  return { tag, attrs, children };
}

function serializeBackground(bg: { type?: string; color?: string }): ParsedNode {
  const bgPrChildren: ParsedNode[] = [];

  if (bg.type === 'solid' && bg.color) {
    bgPrChildren.push({
      tag: 'a:solidFill',
      attrs: {} as Record<string, string>,
      children: [
        { tag: 'a:srgbClr', attrs: { val: bg.color }, children: [] },
      ],
    });
  }

  return {
    tag: 'p:bg',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:bgPr',
        attrs: {} as Record<string, string>,
        children: bgPrChildren,
      },
    ],
  };
}

export function serializeMasterRels(master: SlideMaster, index: number): string {
  const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

  const children: ParsedNode[] = master.layouts.map((layout, i) => ({
    tag: 'Relationship',
    attrs: {
      Id: `rId${i + 1}`,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
      Target: `../slideLayouts/slideLayout${layout.id}.xml`,
    },
    children: [],
  }));

  // Add theme relationship
  children.push({
    tag: 'Relationship',
    attrs: {
      Id: `rId${children.length + 1}`,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      Target: '../theme/theme1.xml',
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
