import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { SheetImage } from '../types.js';

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function serializeDrawing(images: SheetImage[]): string {
  const anchors = images.map((img, i) => {
    const fromChildren: ParsedNode[] = [
      { tag: 'xdr:col', attrs: {}, children: [String(img.position.from.col)] },
      { tag: 'xdr:colOff', attrs: {}, children: [String(img.position.from.colOff || 0)] },
      { tag: 'xdr:row', attrs: {}, children: [String(img.position.from.row)] },
      { tag: 'xdr:rowOff', attrs: {}, children: [String(img.position.from.rowOff || 0)] },
    ];

    const toChildren: ParsedNode[] = [
      { tag: 'xdr:col', attrs: {}, children: [String(img.position.to.col)] },
      { tag: 'xdr:colOff', attrs: {}, children: [String(img.position.to.colOff || 0)] },
      { tag: 'xdr:row', attrs: {}, children: [String(img.position.to.row)] },
      { tag: 'xdr:rowOff', attrs: {}, children: [String(img.position.to.rowOff || 0)] },
    ];

    const nvPicPrChildren: ParsedNode[] = [
      { tag: 'xdr:cNvPr', attrs: { id: String(i + 1), name: img.name || `Picture ${i + 1}` }, children: [] },
      { tag: 'xdr:cNvPicPr', attrs: {}, children: [] },
    ];

    const spPrChildren: ParsedNode[] = [];
    if (img.size) {
      spPrChildren.push({
        tag: 'a:xfrm',
        attrs: {},
        children: [
          { tag: 'a:off', attrs: { x: '0', y: '0' }, children: [] },
          { tag: 'a:ext', attrs: { cx: String(img.size.width), cy: String(img.size.height) }, children: [] },
        ],
      });
    }

    return {
      tag: 'xdr:twoCellAnchor',
      attrs: {},
      children: [
        { tag: 'xdr:from', attrs: {}, children: fromChildren },
        { tag: 'xdr:to', attrs: {}, children: toChildren },
        {
          tag: 'xdr:pic',
          attrs: {},
          children: [
            { tag: 'xdr:nvPicPr', attrs: {}, children: nvPicPrChildren },
            {
              tag: 'xdr:blipFill',
              attrs: {},
              children: [
                { tag: 'a:blip', attrs: { 'r:embed': img.relationshipId }, children: [] },
              ],
            },
            { tag: 'xdr:spPr', attrs: {}, children: spPrChildren },
          ],
        },
        { tag: 'xdr:clientData', attrs: {}, children: [] },
      ],
    };
  });

  const root: ParsedNode = {
    tag: 'xdr:wsDr',
    attrs: { 'xmlns:xdr': XDR_NS, 'xmlns:a': A_NS, 'xmlns:r': R_NS },
    children: anchors,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
