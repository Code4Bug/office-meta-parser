import type { ParsedNode } from '../../core/types.js';
import type {
  TextShape, ImageShape, SlideElement, Position, Hyperlink,
  BodyProperties, ListStyle, ListLevelProperties,
} from '../types.js';
import type { Paragraph, TextRun, ParagraphProperties, RunProperties } from '../../docx/types.js';
import { findChild, findChildAny, getTextContent } from './utils.js';
import { parseShapeStyle, parseColorValue } from './style.js';

export function parseShape(node: ParsedNode): SlideElement | null {
  const txBody = findChildAny(node, 'txBody');
  if (txBody) {
    return parseTextShape(node, txBody);
  }

  const pic = findChildAny(node, 'pic');
  if (pic) {
    return parseImageShape(pic);
  }

  return null;
}

export function parseTextShape(spNode: ParsedNode, txBody: ParsedNode): TextShape {
  const paragraphs: Paragraph[] = [];

  for (const child of txBody.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:p') {
      paragraphs.push(parseParagraph(child));
    }
  }

  const content = paragraphs
    .map(p => p.runs.map(r => r.text).join(''))
    .join('\n');

  const spPr = findChildAny(spNode, 'spPr');
  const position = spPr ? parsePosition(spPr) : { x: 0, y: 0, width: 0, height: 0 };

  let placeholder: { type: string; index?: number } | undefined;
  if (spPr) {
    const ph = findChildAny(spPr, 'ph');
    if (ph) {
      placeholder = {
        type: ph.attrs['type'] || 'body',
      };
      if (ph.attrs['idx']) {
        placeholder.index = parseInt(ph.attrs['idx'], 10);
      }
    }
  }

  const style = spPr ? parseShapeStyle(spPr) : undefined;
  const hyperlink = spPr ? parseHyperlink(spPr) : undefined;

  // Parse preset geometry
  let presetGeom: string | undefined;
  if (spPr) {
    const prstGeom = findChildAny(spPr, 'prstGeom');
    if (prstGeom && prstGeom.attrs['prst']) {
      presetGeom = prstGeom.attrs['prst'];
    }
  }

  // Parse body properties
  const bodyProperties = parseBodyProperties(txBody);

  // Parse list style
  const listStyle = parseListStyle(txBody);

  // Parse rotation from xfrm
  let rotation: number | undefined;
  if (spPr) {
    const xfrm = findChildAny(spPr, 'xfrm');
    if (xfrm && xfrm.attrs['rot']) {
      rotation = parseInt(xfrm.attrs['rot'], 10) / 60000;
    }
  }

  const result: TextShape = {
    type: 'text',
    content,
    position,
    paragraphs,
  };
  if (placeholder) result.placeholder = placeholder;
  if (style) result.style = style;
  if (hyperlink) result.hyperlink = hyperlink;
  if (presetGeom) result.presetGeom = presetGeom;
  if (bodyProperties) result.bodyProperties = bodyProperties;
  if (listStyle) result.listStyle = listStyle;
  if (rotation !== undefined) result.rotation = rotation;
  return result;
}

function parseBodyProperties(txBody: ParsedNode): BodyProperties | undefined {
  const bodyPr = findChildAny(txBody, 'bodyPr');
  if (!bodyPr) return undefined;

  const bp: BodyProperties = {};
  if (bodyPr.attrs['anchor']) bp.anchor = bodyPr.attrs['anchor'] as BodyProperties['anchor'];
  if (bodyPr.attrs['wrap']) bp.wrap = bodyPr.attrs['wrap'] as BodyProperties['wrap'];
  if (bodyPr.attrs['lIns']) bp.leftInset = parseInt(bodyPr.attrs['lIns'], 10) / 914400;
  if (bodyPr.attrs['tIns']) bp.topInset = parseInt(bodyPr.attrs['tIns'], 10) / 914400;
  if (bodyPr.attrs['rIns']) bp.rightInset = parseInt(bodyPr.attrs['rIns'], 10) / 914400;
  if (bodyPr.attrs['bIns']) bp.bottomInset = parseInt(bodyPr.attrs['bIns'], 10) / 914400;
  if (bodyPr.attrs['vert']) bp.vertical = bodyPr.attrs['vert'] as BodyProperties['vertical'];

  // autoFit: normAutofit or spAutoFit
  const normAutofit = findChildAny(bodyPr, 'normAutofit');
  if (normAutofit) bp.autoFit = 'normal';
  const spAutoFit = findChildAny(bodyPr, 'spAutoFit');
  if (spAutoFit) bp.autoFit = 'shape';

  return Object.keys(bp).length > 0 ? bp : undefined;
}

function parseListStyle(txBody: ParsedNode): ListStyle | undefined {
  const lstStyle = findChildAny(txBody, 'lstStyle');
  if (!lstStyle) return undefined;

  const defaultParagraphProperties: ListLevelProperties[] = [];
  for (const child of lstStyle.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:lvl1pPr' || child.tag === 'a:lvl2pPr' || child.tag === 'a:lvl3pPr' ||
        child.tag === 'a:lvl4pPr' || child.tag === 'a:lvl5pPr' || child.tag === 'a:lvl6pPr' ||
        child.tag === 'a:lvl7pPr' || child.tag === 'a:lvl8pPr' || child.tag === 'a:lvl9pPr') {
      const level = parseInt(child.tag.replace('a:lvl', '').replace('pPr', ''), 10);
      const props = parseListLevelProps(child, level);
      if (props) defaultParagraphProperties.push(props);
    }
  }

  return defaultParagraphProperties.length > 0 ? { defaultParagraphProperties } : undefined;
}

function parseListLevelProps(node: ParsedNode, level: number): ListLevelProperties | undefined {
  const props: ListLevelProperties = { level };
  if (node.attrs['algn']) props.alignment = node.attrs['algn'];
  if (node.attrs['marL']) props.marL = parseInt(node.attrs['marL'], 10) / 12700;
  if (node.attrs['indent']) props.indent = parseInt(node.attrs['indent'], 10) / 12700;

  const defRPr = findChildAny(node, 'defRPr');
  if (defRPr && defRPr.attrs['sz']) {
    props.fontScale = parseInt(defRPr.attrs['sz'], 10) / 100;
  }

  return Object.keys(props).length > 1 ? props : undefined;
}

function parseParagraph(node: ParsedNode): Paragraph {
  const runs: TextRun[] = [];
  let properties: ParagraphProperties | undefined;
  let numbering: { level: number; numId: string; format?: string; text?: string } | undefined;

  const pPr = findChildAny(node, 'pPr');
  if (pPr) {
    properties = parseParagraphProperties(pPr);

    // Parse bullet info
    const buNone = findChildAny(pPr, 'buNone');
    if (!buNone) {
      const buChar = findChildAny(pPr, 'buChar');
      const buAutoNum = findChildAny(pPr, 'buAutoNum');
      if (buChar || buAutoNum) {
        const lvl = pPr.attrs['lvl'] ? parseInt(pPr.attrs['lvl'], 10) : 0;
        numbering = {
          level: lvl,
          numId: buAutoNum ? 'buAutoNum' : 'buChar',
        };
        if (buChar && buChar.attrs['char']) numbering.format = buChar.attrs['char'];
        if (buAutoNum && buAutoNum.attrs['type']) numbering.format = buAutoNum.attrs['type'];
        if (buAutoNum && buAutoNum.attrs['startAt']) numbering.text = buAutoNum.attrs['startAt'];
      }
    }
  }

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:r') {
      runs.push(parseRun(child));
    } else if (child.tag === 'a:br') {
      // Line break = empty run
      runs.push({ text: '\n' });
    }
  }

  const para: Paragraph = { type: 'paragraph', runs };
  if (properties) para.properties = properties;
  if (numbering) para.numbering = numbering;
  return para;
}

function parseParagraphProperties(pPr: ParsedNode): ParagraphProperties | undefined {
  const props: ParagraphProperties = {};

  // Alignment
  if (pPr.attrs['algn']) {
    const algn = pPr.attrs['algn'];
    if (algn === 'l') props.alignment = 'left';
    else if (algn === 'ctr') props.alignment = 'center';
    else if (algn === 'r') props.alignment = 'right';
    else if (algn === 'just') props.alignment = 'justify';
  }

  // Indentation
  const indent: { left?: number; right?: number; firstLine?: number } = {};
  if (pPr.attrs['marL']) indent.left = parseInt(pPr.attrs['marL'], 10) / 12700;
  if (pPr.attrs['marR']) indent.right = parseInt(pPr.attrs['marR'], 10) / 12700;
  if (pPr.attrs['indent']) indent.firstLine = parseInt(pPr.attrs['indent'], 10) / 12700;
  if (Object.keys(indent).length > 0) props.indent = indent;

  // Spacing
  const spacing: { before?: number; after?: number; line?: number; lineRule?: 'auto' | 'exact' | 'atLeast' } = {};
  const spcBef = findChildAny(pPr, 'spcBef');
  if (spcBef) {
    const spcPts = findChildAny(spcBef, 'spcPts');
    if (spcPts && spcPts.attrs['val']) spacing.before = parseInt(spcPts.attrs['val'], 10) / 100;
  }
  const spcAft = findChildAny(pPr, 'spcAft');
  if (spcAft) {
    const spcPts = findChildAny(spcAft, 'spcPts');
    if (spcPts && spcPts.attrs['val']) spacing.after = parseInt(spcPts.attrs['val'], 10) / 100;
  }
  const lnSpc = findChildAny(pPr, 'lnSpc');
  if (lnSpc) {
    const spcPct = findChildAny(lnSpc, 'spcPct');
    if (spcPct && spcPct.attrs['val']) {
      spacing.line = parseInt(spcPct.attrs['val'], 10) / 1000;
      spacing.lineRule = 'auto';
    }
    const spcPts = findChildAny(lnSpc, 'spcPts');
    if (spcPts && spcPts.attrs['val']) {
      spacing.line = parseInt(spcPts.attrs['val'], 10) / 100;
      spacing.lineRule = 'exact';
    }
  }
  if (Object.keys(spacing).length > 0) props.spacing = spacing;

  return Object.keys(props).length > 0 ? props : undefined;
}

function parseRun(node: ParsedNode): TextRun {
  let text = '';

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'a:t') {
      text = getTextContent(child);
    }
  }

  const rPr = findChildAny(node, 'rPr');
  const run: TextRun = { text };
  if (rPr) {
    if (rPr.attrs['b'] === '1') run.bold = true;
    if (rPr.attrs['i'] === '1') run.italic = true;
    if (rPr.attrs['u'] && rPr.attrs['u'] !== 'none') run.underline = true;
    if (rPr.attrs['strike'] && rPr.attrs['strike'] !== 'noStrike') run.strike = true;
    if (rPr.attrs['sz']) run.fontSize = parseInt(rPr.attrs['sz'], 10) / 100;
    if (rPr.attrs['cap']) run.caps = rPr.attrs['cap'] === 'all';
    if (rPr.attrs['spc']) run.characterSpacing = parseInt(rPr.attrs['spc'], 10) / 100;

    const solidFill = findChildAny(rPr, 'solidFill');
    if (solidFill) {
      const color = parseColorValue(solidFill);
      if (color) run.color = color;
    }

    const latin = findChildAny(rPr, 'latin');
    if (latin && latin.attrs['typeface']) run.fontFamily = latin.attrs['typeface'];

    if (rPr.attrs['baseline']) {
      const baseline = parseInt(rPr.attrs['baseline'], 10);
      if (baseline > 0) run.superscript = true;
      else if (baseline < 0) run.subscript = true;
    }
  }

  return run;
}

export function parseImageShape(pic: ParsedNode): ImageShape {
  const blipFill = findChildAny(pic, 'blipFill');
  let relationshipId = '';

  if (blipFill) {
    const blip = findChildAny(blipFill, 'blip');
    if (blip) {
      relationshipId = blip.attrs['r:embed'] || '';
    }
  }

  const spPr = findChildAny(pic, 'spPr');
  const position = spPr ? parsePosition(spPr) : { x: 0, y: 0, width: 0, height: 0 };
  const style = spPr ? parseShapeStyle(spPr) : undefined;
  const hyperlink = spPr ? parseHyperlink(spPr) : undefined;

  // Parse rotation
  let rotation: number | undefined;
  if (spPr) {
    const xfrm = findChildAny(spPr, 'xfrm');
    if (xfrm && xfrm.attrs['rot']) {
      rotation = parseInt(xfrm.attrs['rot'], 10) / 60000;
    }
  }

  const result: ImageShape = {
    type: 'image',
    relationshipId,
    position,
  };
  if (style) result.style = style;
  if (hyperlink) result.hyperlink = hyperlink;
  if (rotation !== undefined) result.rotation = rotation;
  return result;
}

function parsePosition(spPr: ParsedNode): Position {
  const xfrm = findChildAny(spPr, 'xfrm');
  if (!xfrm) return { x: 0, y: 0, width: 0, height: 0 };

  const off = findChildAny(xfrm, 'off');
  const ext = findChildAny(xfrm, 'ext');

  const pos: Position = {
    x: off?.attrs['x'] ? parseInt(off.attrs['x'], 10) : 0,
    y: off?.attrs['y'] ? parseInt(off.attrs['y'], 10) : 0,
    width: ext?.attrs['cx'] ? parseInt(ext.attrs['cx'], 10) : 0,
    height: ext?.attrs['cy'] ? parseInt(ext.attrs['cy'], 10) : 0,
  };

  // Parse flip
  if (xfrm.attrs['flipH'] === '1') pos.flipH = true;
  if (xfrm.attrs['flipV'] === '1') pos.flipV = true;
  if (xfrm.attrs['rot']) {
    pos.rotation = parseInt(xfrm.attrs['rot'], 10) / 60000;
  }

  return pos;
}

function parseHyperlink(spPr: ParsedNode): Hyperlink | undefined {
  const hlinkClick = findChildAny(spPr, 'hlinkClick');
  if (!hlinkClick) return undefined;

  const url = hlinkClick.attrs['r:id'] || hlinkClick.attrs['action'] || '';
  if (!url) return undefined;

  const tooltip = hlinkClick.attrs['tooltip'];
  const result: Hyperlink = { url };
  if (tooltip) result.tooltip = tooltip;
  return result;
}
