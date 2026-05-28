import type { RawDocument, ParsedNode } from '../../core/types.js';
import type {
  SlideMaster, SlideLayout, Placeholder, Position,
  MasterTxStyles, TextStyle, TextLevelStyle, FillStyle,
} from '../types.js';
import { findChild, findChildAny } from './utils.js';
import { parseColorValue } from './style.js';

export function extractMasters(raw: RawDocument): SlideMaster[] {
  const masters: SlideMaster[] = [];

  for (const [path, xml] of raw.parts) {
    if (path.startsWith('ppt/slideMasters/slideMaster') && path.endsWith('.xml')) {
      masters.push(parseMaster(raw, xml, path));
    }
  }

  return masters;
}

export function extractLayouts(raw: RawDocument): SlideLayout[] {
  const layouts: SlideLayout[] = [];

  for (const [path, xml] of raw.parts) {
    if (path.startsWith('ppt/slideLayouts/slideLayout') && path.endsWith('.xml')) {
      layouts.push(parseLayout(xml, path));
    }
  }

  return layouts;
}

function parseMaster(raw: RawDocument, node: ParsedNode, path: string): SlideMaster {
  const id = path.match(/slideMaster(\d+)\.xml/)?.[1] || '';

  // Parse background
  const cSld = findChild(node, 'p:cSld');
  const bg = cSld ? findChild(cSld, 'p:bg') : undefined;
  const background = bg ? parseBackground(bg) : undefined;

  // Parse associated layouts via relationships
  const relsPath = path.replace('ppt/slideMasters/', 'ppt/slideMasters/_rels/') + '.rels';
  const rels = raw.rels.get(relsPath) || [];
  const layouts: SlideLayout[] = [];

  for (const rel of rels) {
    if (rel.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout') {
      const layoutPath = resolveRelativePath('ppt/slideMasters/', rel.target);
      const layoutXml = raw.parts.get(layoutPath);
      if (layoutXml) {
        layouts.push(parseLayout(layoutXml, layoutPath));
      }
    }
  }

  const master: SlideMaster = { id, layouts };
  if (background) master.background = background;

  // Parse txStyles
  const txStyles = parseTxStyles(node);
  if (txStyles) master.txStyles = txStyles;

  return master;
}

function parseTxStyles(node: ParsedNode): MasterTxStyles | undefined {
  const txStyles = findChild(node, 'p:txStyles');
  if (!txStyles) return undefined;

  const styles: MasterTxStyles = {};

  const titleStyle = findChild(txStyles, 'p:titleStyle');
  if (titleStyle) styles.titleStyle = parseTextStyle(titleStyle);

  const bodyStyle = findChild(txStyles, 'p:bodyStyle');
  if (bodyStyle) styles.bodyStyle = parseTextStyle(bodyStyle);

  const otherStyle = findChild(txStyles, 'p:otherStyle');
  if (otherStyle) styles.otherStyle = parseTextStyle(otherStyle);

  return Object.keys(styles).length > 0 ? styles : undefined;
}

function parseTextStyle(node: ParsedNode): TextStyle {
  const levels: TextLevelStyle[] = [];

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:lvl1pPr' || child.tag === 'a:lvl2pPr' || child.tag === 'a:lvl3pPr' ||
        child.tag === 'a:lvl4pPr' || child.tag === 'a:lvl5pPr' || child.tag === 'a:lvl6pPr' ||
        child.tag === 'a:lvl7pPr' || child.tag === 'a:lvl8pPr' || child.tag === 'a:lvl9pPr') {
      const level = parseInt(child.tag.replace('a:lvl', '').replace('pPr', ''), 10);
      const props = parseTextLevelStyle(child, level);
      if (props) levels.push(props);
    }
  }

  return { levels };
}

function parseTextLevelStyle(node: ParsedNode, level: number): TextLevelStyle | undefined {
  const props: TextLevelStyle = { level };

  if (node.attrs['algn']) props.alignment = node.attrs['algn'];
  if (node.attrs['marL']) props.marL = parseInt(node.attrs['marL'], 10) / 12700;
  if (node.attrs['indent']) props.indent = parseInt(node.attrs['indent'], 10) / 12700;

  // Spacing
  const spcBef = findChildAny(node, 'spcBef');
  if (spcBef) {
    const spcPts = findChildAny(spcBef, 'spcPts');
    if (spcPts && spcPts.attrs['val']) props.spcBef = parseInt(spcPts.attrs['val'], 10) / 100;
  }
  const spcAft = findChildAny(node, 'spcAft');
  if (spcAft) {
    const spcPts = findChildAny(spcAft, 'spcPts');
    if (spcPts && spcPts.attrs['val']) props.spcAft = parseInt(spcPts.attrs['val'], 10) / 100;
  }
  const lnSpc = findChildAny(node, 'lnSpc');
  if (lnSpc) {
    const spcPct = findChildAny(lnSpc, 'spcPct');
    if (spcPct && spcPct.attrs['val']) props.lnSpc = parseInt(spcPct.attrs['val'], 10) / 1000;
  }

  // Font scale from defRPr
  const defRPr = findChildAny(node, 'defRPr');
  if (defRPr && defRPr.attrs['sz']) {
    props.fontScale = parseInt(defRPr.attrs['sz'], 10) / 100;
  }

  return Object.keys(props).length > 1 ? props : undefined;
}

function parseLayout(node: ParsedNode, path: string): SlideLayout {
  const id = path.match(/slideLayout(\d+)\.xml/)?.[1] || '';

  const cSld = findChild(node, 'p:cSld');
  const spTree = cSld ? findChild(cSld, 'p:spTree') : undefined;

  const placeholders: Placeholder[] = [];
  if (spTree) {
    for (const child of spTree.children) {
      if (typeof child === 'string' || child.tag !== 'p:sp') continue;
      const spPr = findChild(child, 'p:spPr');
      if (!spPr) continue;
      const ph = findChild(spPr, 'p:ph');
      if (!ph) continue;

      const position = parsePosition(spPr);
      const placeholder: Placeholder = {
        type: ph.attrs['type'] || 'body',
        position,
      };
      if (ph.attrs['idx']) {
        placeholder.index = parseInt(ph.attrs['idx'], 10);
      }
      placeholders.push(placeholder);
    }
  }

  const name = cSld?.attrs['name'] || undefined;

  const layout: SlideLayout = { id, placeholders };
  if (name) layout.name = name;
  return layout;
}

function parseBackground(bg: ParsedNode): FillStyle | undefined {
  const bgPr = findChild(bg, 'p:bgPr');
  if (!bgPr) return undefined;

  const solidFill = findChild(bgPr, 'a:solidFill');
  if (solidFill) {
    const srgbClr = findChild(solidFill, 'a:srgbClr');
    if (srgbClr && srgbClr.attrs['val']) {
      return { type: 'solid' as const, color: srgbClr.attrs['val'] };
    }
    const schemeClr = findChild(solidFill, 'a:schemeClr');
    if (schemeClr && schemeClr.attrs['val']) {
      return { type: 'solid' as const, color: `scheme:${schemeClr.attrs['val']}` };
    }
  }

  return undefined;
}

function parsePosition(spPr: ParsedNode): Position {
  const xfrm = findChild(spPr, 'a:xfrm');
  if (!xfrm) return { x: 0, y: 0, width: 0, height: 0 };

  const off = findChild(xfrm, 'a:off');
  const ext = findChild(xfrm, 'a:ext');

  return {
    x: off?.attrs['x'] ? parseInt(off.attrs['x'], 10) : 0,
    y: off?.attrs['y'] ? parseInt(off.attrs['y'], 10) : 0,
    width: ext?.attrs['cx'] ? parseInt(ext.attrs['cx'], 10) : 0,
    height: ext?.attrs['cy'] ? parseInt(ext.attrs['cy'], 10) : 0,
  };
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
