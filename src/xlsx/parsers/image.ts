import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { SheetImage } from '../types.js';
import { findChild } from './utils.js';

export function extractImages(raw: RawDocument, sheetPath: string): SheetImage[] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  // Find drawing relationship in worksheet rels
  const sheetDir = sheetPath.substring(0, sheetPath.lastIndexOf('/'));
  const sheetName = sheetPath.substring(sheetPath.lastIndexOf('/') + 1);
  const relsPath = `${sheetDir}/_rels/${sheetName}.rels`;
  const rels = raw.rels.get(relsPath) || [];

  const drawingRel = rels.find(r =>
    r.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
  );

  if (!drawingRel) return [];

  // Resolve drawing path
  const drawingPath = resolveRelativePath(sheetDir + '/', drawingRel.target);
  const drawingXml = raw.parts.get(drawingPath);
  if (!drawingXml) return [];

  // Parse drawing XML
  const images: SheetImage[] = [];
  const twoCellAnchors = findChildren(drawingXml, 'xdr:twoCellAnchor');
  const oneCellAnchors = findChildren(drawingXml, 'xdr:oneCellAnchor');

  for (const anchor of [...twoCellAnchors, ...oneCellAnchors]) {
    const pic = findChild(anchor, 'xdr:pic');
    if (!pic) continue;

    const nvPicPr = findChild(pic, 'xdr:nvPicPr');
    const cNvPr = nvPicPr ? findChild(nvPicPr, 'xdr:cNvPr') : undefined;
    const blipFill = findChild(pic, 'xdr:blipFill');
    const blip = blipFill ? findChild(blipFill, 'a:blip') : undefined;

    if (!blip) continue;

    // Parse position
    const from = findChild(anchor, 'xdr:from');
    const to = findChild(anchor, 'xdr:to');
    let position: SheetImage['position'];
    let size: SheetImage['size'] | undefined;

    if (from) {
      position = {
        from: parseAnchorPoint(from),
        to: to ? parseAnchorPoint(to) : { col: 0, row: 0 },
      };
    } else {
      // oneCellAnchor has ext instead of to
      const ext = findChild(anchor, 'xdr:ext');
      if (ext) {
        position = {
          from: { col: 0, row: 0 },
          to: { col: 0, row: 0 },
        };
        if (ext.attrs['cx']) size = { width: parseInt(ext.attrs['cx'], 10), height: parseInt(ext.attrs['cy'] || '0', 10) };
      } else {
        position = { from: { col: 0, row: 0 }, to: { col: 0, row: 0 } };
      }
    }

    const image: SheetImage = {
      relationshipId: blip.attrs['r:embed'] || '',
      position,
    };

    if (size) image.size = size;

    if (cNvPr) {
      if (cNvPr.attrs['name']) image.name = cNvPr.attrs['name'];
      if (cNvPr.attrs['descr']) image.description = cNvPr.attrs['descr'];
    }

    images.push(image);
  }

  return images;
}

function parseAnchorPoint(node: ParsedNode): { col: number; row: number; colOff?: number; rowOff?: number } {
  const col = findChild(node, 'xdr:col');
  const row = findChild(node, 'xdr:row');
  const colOff = findChild(node, 'xdr:colOff');
  const rowOff = findChild(node, 'xdr:rowOff');

  const result: { col: number; row: number; colOff?: number; rowOff?: number } = {
    col: col && col.children.length > 0 && typeof col.children[0] === 'string' ? parseInt(col.children[0], 10) : 0,
    row: row && row.children.length > 0 && typeof row.children[0] === 'string' ? parseInt(row.children[0], 10) : 0,
  };

  if (colOff && colOff.children.length > 0 && typeof colOff.children[0] === 'string') {
    result.colOff = parseInt(colOff.children[0], 10);
  }
  if (rowOff && rowOff.children.length > 0 && typeof rowOff.children[0] === 'string') {
    result.rowOff = parseInt(rowOff.children[0], 10);
  }

  return result;
}

function findChildren(node: ParsedNode, tag: string): ParsedNode[] {
  return node.children.filter((c): c is ParsedNode => typeof c !== 'string' && c.tag === tag);
}

function resolveRelativePath(base: string, relative: string): string {
  const baseParts = base.split('/').filter(Boolean);
  const relParts = relative.split('/').filter(Boolean);
  const result = [...baseParts];
  for (const part of relParts) {
    if (part === '..') {
      result.pop();
    } else if (part !== '.') {
      result.push(part);
    }
  }
  return result.join('/');
}
