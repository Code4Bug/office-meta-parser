import type { ParsedNode, RawDocument } from '../../core/types.js';
import type { Image } from '../types.js';
import { findChild } from './utils.js';

export function parseImage(raw: RawDocument, node: ParsedNode): Image | null {
  // w:drawing → wp:inline or wp:anchor
  const inline = findChild(node, 'wp:inline');
  const anchor = findChild(node, 'wp:anchor');
  const wrapper = inline || anchor;
  if (!wrapper) return null;

  const isFloating = !!anchor;

  // Extract relationshipId from a:blip
  const graphic = findChild(wrapper, 'a:graphic');
  if (!graphic) return null;
  const graphicData = findChild(graphic, 'a:graphicData');
  if (!graphicData) return null;
  const pic = findChild(graphicData, 'pic:pic');
  if (!pic) return null;
  const blipFill = findChild(pic, 'pic:blipFill');
  if (!blipFill) return null;
  const blip = findChild(blipFill, 'a:blip');
  if (!blip) return null;

  const relationshipId = blip.attrs['r:embed'] || '';

  // Extract size from wp:extent
  const extent = findChild(wrapper, 'wp:extent');
  const width = extent?.attrs['cx'] ? parseInt(extent.attrs['cx'], 10) : undefined;
  const height = extent?.attrs['cy'] ? parseInt(extent.attrs['cy'], 10) : undefined;

  // Extract alt text
  const docPr = findChild(wrapper, 'wp:docPr');
  const alt = docPr?.attrs['descr'] || docPr?.attrs['name'] || undefined;

  const image: Image = { type: 'image', relationshipId, width, height, alt };

  if (isFloating) {
    image.isFloating = true;

    // Wrap type
    for (const child of anchor!.children) {
      if (typeof child === 'string') continue;
      if (child.tag === 'wp:wrapSquare') { image.wrapType = 'square'; break; }
      if (child.tag === 'wp:wrapTight') { image.wrapType = 'tight'; break; }
      if (child.tag === 'wp:wrapThrough') { image.wrapType = 'through'; break; }
      if (child.tag === 'wp:wrapTopBottom') { image.wrapType = 'topBottom'; break; }
      if (child.tag === 'wp:wrapNone') { image.wrapType = 'none'; break; }
      if (child.tag === 'wp:wrapBehindText') { image.wrapType = 'behind'; break; }
      if (child.tag === 'wp:wrapInFrontOfText') { image.wrapType = 'front'; break; }
    }

    // Position offsets (EMU)
    const posH = findChild(anchor!, 'wp:positionH');
    if (posH) {
      const offset = findChild(posH, 'wp:posOffset');
      if (offset) {
        const text = offset.children.find(c => typeof c === 'string');
        if (text) image.posX = parseInt(text as string, 10);
      }
    }
    const posV = findChild(anchor!, 'wp:positionV');
    if (posV) {
      const offset = findChild(posV, 'wp:posOffset');
      if (offset) {
        const text = offset.children.find(c => typeof c === 'string');
        if (text) image.posY = parseInt(text as string, 10);
      }
    }
  }

  return image;
}
