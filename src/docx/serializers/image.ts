import type { ParsedNode } from '../../core/types.js';
import type { Image } from '../types.js';

export function serializeImageParagraph(image: Image): ParsedNode {
  const cx = image.width ? String(image.width) : '914400';
  const cy = image.height ? String(image.height) : '914400';
  const graphic = buildGraphicNode(image.relationshipId, cx, cy);

  const wrapper: ParsedNode = image.isFloating
    ? buildAnchorNode(image, cx, cy, graphic)
    : buildInlineNode(cx, cy, graphic, image.alt);

  return {
    tag: 'w:p',
    attrs: {},
    children: [
      { tag: 'w:r', attrs: {}, children: [{ tag: 'w:drawing', attrs: {}, children: [wrapper] }] },
    ],
  };
}

function buildGraphicNode(relId: string, cx: string, cy: string): ParsedNode {
  return {
    tag: 'a:graphic',
    attrs: {},
    children: [
      {
        tag: 'a:graphicData',
        attrs: { uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture' },
        children: [
          {
            tag: 'pic:pic',
            attrs: {},
            children: [
              {
                tag: 'pic:blipFill',
                attrs: {},
                children: [{ tag: 'a:blip', attrs: { 'r:embed': relId }, children: [] }],
              },
              {
                tag: 'pic:spPr',
                attrs: {},
                children: [
                  {
                    tag: 'a:xfrm',
                    attrs: {},
                    children: [{ tag: 'a:ext', attrs: { cx, cy }, children: [] }],
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

function buildInlineNode(cx: string, cy: string, graphic: ParsedNode, alt?: string): ParsedNode {
  const children: ParsedNode[] = [
    { tag: 'wp:extent', attrs: { cx, cy }, children: [] },
  ];
  if (alt) {
    children.push({ tag: 'wp:docPr', attrs: { descr: alt }, children: [] });
  }
  children.push(graphic);
  return {
    tag: 'wp:inline',
    attrs: { distT: '0', distB: '0', distL: '0', distR: '0' },
    children,
  };
}

function buildAnchorNode(image: Image, cx: string, cy: string, graphic: ParsedNode): ParsedNode {
  const wrapTag = getWrapTag(image.wrapType);
  const anchorChildren: ParsedNode[] = [
    { tag: 'wp:simplePos', attrs: { x: '0', y: '0' }, children: [] },
    {
      tag: 'wp:positionH',
      attrs: { relativeFrom: 'column' },
      children: [
        { tag: 'wp:posOffset', attrs: {}, children: [String(image.posX ?? 0)] },
      ],
    },
    {
      tag: 'wp:positionV',
      attrs: { relativeFrom: 'paragraph' },
      children: [
        { tag: 'wp:posOffset', attrs: {}, children: [String(image.posY ?? 0)] },
      ],
    },
    { tag: 'wp:extent', attrs: { cx, cy }, children: [] },
    { tag: 'wp:effectExtent', attrs: { l: '0', t: '0', r: '0', b: '0' }, children: [] },
    { tag: wrapTag, attrs: { wrapText: 'bothSides' }, children: [] },
  ];
  if (image.alt) {
    anchorChildren.push({ tag: 'wp:docPr', attrs: { descr: image.alt }, children: [] });
  }
  anchorChildren.push(graphic);

  return {
    tag: 'wp:anchor',
    attrs: {
      distT: '0', distB: '114300', distL: '114300', distR: '114300',
      simplePos: '0', relativeHeight: '251659264', behindDoc: '0',
      locked: '0', layoutInCell: '1', allowOverlap: '1',
    },
    children: anchorChildren,
  };
}

function getWrapTag(wrapType?: Image['wrapType']): string {
  switch (wrapType) {
    case 'tight': return 'wp:wrapTight';
    case 'through': return 'wp:wrapThrough';
    case 'topBottom': return 'wp:wrapTopBottom';
    case 'none': return 'wp:wrapNone';
    case 'behind': return 'wp:wrapBehindText';
    case 'front': return 'wp:wrapInFrontOfText';
    default: return 'wp:wrapSquare';
  }
}
