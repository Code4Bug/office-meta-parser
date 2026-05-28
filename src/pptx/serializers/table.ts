import type { ParsedNode } from '../../core/types.js';
import type { TableShape, FillStyle, BorderStyle } from '../types.js';

export function serializeTableShape(shape: TableShape, nextId?: { value: number }): ParsedNode {
  const tblChildren: ParsedNode[] = [];

  for (const row of shape.rows) {
    const trAttrs: Record<string, string> = {};
    if (row.height) trAttrs['h'] = String(row.height);

    const tcChildren: ParsedNode[] = [];
    for (const cell of row.cells) {
      const cellAttrs: Record<string, string> = {};
      if (cell.width) cellAttrs['w'] = String(cell.width);
      if (cell.colspan) cellAttrs['gridSpan'] = String(cell.colspan);
      if (cell.rowspan) cellAttrs['rowSpan'] = String(cell.rowspan);

      const txBodyChildren: ParsedNode[] = cell.content.map(para => ({
        tag: 'a:p',
        attrs: {} as Record<string, string>,
        children: para.runs.map(run => ({
          tag: 'a:r',
          attrs: {} as Record<string, string>,
          children: [
            { tag: 'a:rPr', attrs: { lang: 'en-US' }, children: [] },
            { tag: 'a:t', attrs: {} as Record<string, string>, children: [run.text] },
          ],
        })),
      }));

      const tcNodeChildren: ParsedNode[] = [{
        tag: 'a:txBody',
        attrs: {} as Record<string, string>,
        children: txBodyChildren,
      }];

      if (cell.style) {
        const tcPrChildren: ParsedNode[] = [];
        if (cell.style.fill) tcPrChildren.push(...serializeFillNodes(cell.style.fill));
        if (cell.style.borders) {
          const { top, bottom, left, right } = cell.style.borders;
          if (left) tcPrChildren.push(serializeBorderSide('a:lnL', left));
          if (right) tcPrChildren.push(serializeBorderSide('a:lnR', right));
          if (top) tcPrChildren.push(serializeBorderSide('a:lnT', top));
          if (bottom) tcPrChildren.push(serializeBorderSide('a:lnB', bottom));
        }
        tcNodeChildren.push({ tag: 'a:tcPr', attrs: {} as Record<string, string>, children: tcPrChildren });
      }

      tcChildren.push({
        tag: 'a:tc',
        attrs: cellAttrs,
        children: tcNodeChildren,
      });
    }

    tblChildren.push({
      tag: 'a:tr',
      attrs: trAttrs,
      children: tcChildren,
    });
  }

  return {
    tag: 'p:graphicFrame',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:nvGraphicFramePr',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'p:cNvPr', attrs: { id: String(nextId ? nextId.value++ : 3), name: 'Table 1' }, children: [] },
          { tag: 'p:cNvGraphicFramePr', attrs: {} as Record<string, string>, children: [] },
          { tag: 'p:nvPr', attrs: {} as Record<string, string>, children: [] },
        ],
      },
      {
        tag: 'p:xfrm',
        attrs: {} as Record<string, string>,
        children: [
          { tag: 'a:off', attrs: { x: String(shape.position.x), y: String(shape.position.y) }, children: [] },
          { tag: 'a:ext', attrs: { cx: String(shape.position.width), cy: String(shape.position.height) }, children: [] },
        ],
      },
      {
        tag: 'a:graphic',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'a:graphicData',
            attrs: { uri: 'http://schemas.openxmlformats.org/drawingml/2006/table' },
            children: [
              {
                tag: 'a:tbl',
                attrs: {} as Record<string, string>,
                children: [
                  {
                    tag: 'a:tblPr',
                    attrs: shape.tableStyleId ? { styleId: shape.tableStyleId } : {} as Record<string, string>,
                    children: [],
                  },
                  ...tblChildren,
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

function serializeFillNodes(fill: FillStyle): ParsedNode[] {
  if (fill.type === 'solid' && fill.color) {
    return [{
      tag: 'a:solidFill',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: fill.color }, children: [] }],
    }];
  }
  if (fill.type === 'none') {
    return [{ tag: 'a:noFill', attrs: {} as Record<string, string>, children: [] }];
  }
  return [];
}

function serializeBorderSide(tag: string, border: BorderStyle): ParsedNode {
  const attrs: Record<string, string> = {};
  if (border.width !== undefined) attrs.w = String(Math.round(border.width * 12700));
  if (border.compound) attrs.cmpd = border.compound;
  if (border.cap) attrs.cap = border.cap;

  const children: ParsedNode[] = [];
  if (border.color) {
    children.push({
      tag: 'a:solidFill',
      attrs: {} as Record<string, string>,
      children: [{ tag: 'a:srgbClr', attrs: { val: border.color }, children: [] }],
    });
  }
  const dashVal = border.dashType || (border.style === 'dashed' ? 'dash' : border.style === 'dotted' ? 'dot' : undefined);
  if (dashVal) children.push({ tag: 'a:prstDash', attrs: { val: dashVal }, children: [] });

  return { tag, attrs, children };
}
