import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { Slide, SlideElement, GroupShape, Position, Transition, FillStyle, MediaShape, SlideComment } from '../types.js';
import { parseShape, parseImageShape } from './shape.js';
import { parseTableShape } from './table.js';
import { parseAnimations } from './animation.js';
import { parseFill, parseColorValue } from './style.js';
import { findChild, findChildAny, getTextContent } from './utils.js';

export function extractSlides(raw: RawDocument): Slide[] {
  const presentationXml = raw.parts.get('ppt/presentation.xml');
  if (!presentationXml) return [];

  const sldIdLst = findChild(presentationXml, 'p:sldIdLst');
  if (!sldIdLst) return [];

  const slides: Slide[] = [];

  for (const child of sldIdLst.children) {
    if (typeof child === 'string' || child.tag !== 'p:sldId') continue;

    const relId = child.attrs['r:id'] || '';
    const slidePath = findSlidePath(raw, relId);

    if (slidePath) {
      const slideXml = raw.parts.get(slidePath);
      if (slideXml) {
        const slide = parseSlide(slideXml);
        slide.notes = extractNotes(raw, slidePath);
        const comments = extractComments(raw, slidePath);
        if (comments) slide.comments = comments;
        slides.push(slide);
      }
    }
  }

  return slides;
}

export function extractSlideSize(raw: RawDocument): { width: number; height: number } | undefined {
  const presentationXml = raw.parts.get('ppt/presentation.xml');
  if (!presentationXml) return undefined;

  const sldSz = findChild(presentationXml, 'p:sldSz');
  if (!sldSz) return undefined;

  const width = sldSz.attrs['cx'] ? parseInt(sldSz.attrs['cx'], 10) : undefined;
  const height = sldSz.attrs['cy'] ? parseInt(sldSz.attrs['cy'], 10) : undefined;
  if (width && height) return { width, height };
  return undefined;
}

export function extractNotesSize(raw: RawDocument): { width: number; height: number } | undefined {
  const presentationXml = raw.parts.get('ppt/presentation.xml');
  if (!presentationXml) return undefined;

  const notesSz = findChild(presentationXml, 'p:notesSz');
  if (!notesSz) return undefined;

  const width = notesSz.attrs['cx'] ? parseInt(notesSz.attrs['cx'], 10) : undefined;
  const height = notesSz.attrs['cy'] ? parseInt(notesSz.attrs['cy'], 10) : undefined;
  if (width && height) return { width, height };
  return undefined;
}

function findSlidePath(raw: RawDocument, relId: string): string | undefined {
  const presRels = raw.rels.get('ppt/_rels/presentation.xml.rels');
  if (presRels) {
    const rel = presRels.find(r => r.id === relId);
    if (rel) {
      return 'ppt/' + rel.target;
    }
  }
  return undefined;
}

function extractNotes(raw: RawDocument, slidePath: string): string | undefined {
  const slideRelsPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
  const slideRels = raw.rels.get(slideRelsPath);
  if (!slideRels) return undefined;

  const notesRel = slideRels.find(r =>
    r.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
  );
  if (!notesRel) return undefined;

  const notesPath = resolveRelativePath('ppt/slides/', notesRel.target);
  const notesXml = raw.parts.get(notesPath);
  if (!notesXml) return undefined;

  return extractNotesText(notesXml);
}

function extractComments(raw: RawDocument, slidePath: string): SlideComment[] | undefined {
  const slideRelsPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
  const slideRels = raw.rels.get(slideRelsPath);
  if (!slideRels) return undefined;

  const commentsRel = slideRels.find(r =>
    r.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
  );
  if (!commentsRel) return undefined;

  const commentsPath = resolveRelativePath('ppt/slides/', commentsRel.target);
  const commentsXml = raw.parts.get(commentsPath);
  if (!commentsXml) return undefined;

  return parseCommentsXml(commentsXml);
}

function parseCommentsXml(node: ParsedNode): SlideComment[] {
  if (typeof node === 'string') return [];

  // 提取作者列表
  const authorMap = new Map<string, string>();
  const authorLst = findChild(node, 'p:authorLst');
  if (authorLst) {
    for (const child of authorLst.children) {
      if (typeof child === 'string' || child.tag !== 'p:author') continue;
      const authorId = child.attrs['id'] || '0';
      const name = child.attrs['name'] || '';
      authorMap.set(authorId, name);
    }
  }

  const comments: SlideComment[] = [];

  // 找到 p:cmLst 或直接遍历子节点
  const cmList = findChild(node, 'p:cmLst') || node;

  for (const child of (cmList.children || [])) {
    if (typeof child === 'string' || child.tag !== 'p:cm') continue;
    const comment = parseCommentNode(child, authorMap);
    if (comment) comments.push(comment);
  }

  return comments;
}

function parseCommentNode(node: ParsedNode, authorMap: Map<string, string>): SlideComment | undefined {
  if (typeof node === 'string') return undefined;

  const id = node.attrs['id'] || '';
  const authorId = parseInt(node.attrs['authorId'] || '0', 10);
  const dt = node.attrs['dt'] || '';

  // 从 authorMap 查找作者名
  const authorName = authorMap.get(String(authorId)) || '';

  // 提取文本
  let text = '';
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'p:text') {
      text = getTextContent(child);
    }
  }

  // 解析嵌套回复
  const replies: SlideComment[] = [];
  for (const child of node.children) {
    if (typeof child === 'string' || child.tag !== 'p:cm') continue;
    const reply = parseCommentNode(child, authorMap);
    if (reply) replies.push(reply);
  }

  const comment: SlideComment = {
    id,
    authorId,
    authorName,
    text,
  };
  if (dt) comment.date = dt;
  if (replies.length > 0) comment.replies = replies;

  return comment;
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

function extractNotesText(node: ParsedNode): string {
  const texts: string[] = [];
  collectText(node, texts);
  return texts.join('\n');
}

function collectText(node: ParsedNode, texts: string[]): void {
  if (typeof node === 'string') return;
  if (node.tag === 'a:t') {
    const text = getTextContent(node);
    if (text) texts.push(text);
  }
  for (const child of node.children) {
    if (typeof child !== 'string') collectText(child, texts);
  }
}

export function parseSlide(node: ParsedNode): Slide {
  const cSld = findChild(node, 'p:cSld');
  if (!cSld) return { elements: [] };

  const spTree = findChild(cSld, 'p:spTree');
  if (!spTree) return { elements: [] };

  const elements: SlideElement[] = [];

  for (const child of spTree.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'p:sp') {
      const shape = parseShape(child);
      if (shape) elements.push(shape);
    } else if (child.tag === 'p:pic') {
      const media = parseMediaShape(child);
      if (media) {
        elements.push(media);
      } else {
        const shape = parseImageShape(child);
        if (shape) elements.push(shape);
      }
    } else if (child.tag === 'p:grpSp') {
      const group = parseGroupShape(child);
      if (group) elements.push(group);
    } else if (child.tag === 'p:graphicFrame') {
      const table = parseTableShape(child);
      if (table) elements.push(table);
    }
  }

  const transition = parseTransition(node);
  const animations = parseAnimations(node);

  const slide: Slide = { elements };
  if (transition) slide.transition = transition;
  if (animations.length > 0) slide.animations = animations;

  // Parse slide background
  const bg = findChild(cSld, 'p:bg');
  if (bg) {
    slide.background = parseSlideBackground(bg);
  }

  // Parse showMasterSp / showMasterPhAnim
  if (node.attrs['showMasterSp'] === '0') slide.showMasterSp = false;
  if (node.attrs['showMasterPhAnim'] === '0') slide.showMasterPhAnim = false;

  // Parse color map override
  const clrMapOvr = findChild(node, 'p:clrMapOvr');
  if (clrMapOvr) {
    slide.clrMap = parseClrMap(clrMapOvr);
  }

  return slide;
}

function parseSlideBackground(bg: ParsedNode): FillStyle | undefined {
  const bgPr = findChild(bg, 'p:bgPr');
  if (!bgPr) return undefined;

  // Use the shared parseFill from style.ts
  return parseFill(bgPr);
}

function parseClrMap(node: ParsedNode): Record<string, string> | undefined {
  const map: Record<string, string> = {};
  // The clrMapOvr can contain a:masterClrMapping (inherits) or a:overrideClrMapping
  const override = findChild(node, 'a:overrideClrMapping');
  const target = override || node;

  for (const key of ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3',
    'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) {
    if (target.attrs[key]) {
      map[key] = target.attrs[key];
    }
  }

  return Object.keys(map).length > 0 ? map : undefined;
}

function parseTransition(node: ParsedNode): Transition | undefined {
  const transitionNode = findChild(node, 'p:transition');
  if (!transitionNode) return undefined;

  const transition: Transition = { type: 'none' };

  // Parse speed attribute
  if (transitionNode.attrs['spd']) {
    transition.duration = transitionNode.attrs['spd'] === 'fast' ? 500 :
      transitionNode.attrs['spd'] === 'med' ? 1000 : 2000;
  }

  // Parse advClick / advTm
  if (transitionNode.attrs['advClick'] === '0') transition.advClick = false;
  if (transitionNode.attrs['advTm']) {
    transition.advTime = parseInt(transitionNode.attrs['advTm'], 10);
  }

  // Parse transition type from child elements
  for (const child of transitionNode.children) {
    if (typeof child === 'string') continue;
    if (child.tag.startsWith('p:')) {
      const type = child.tag.replace('p:', '');
      if (['blinds', 'checker', 'circle', 'dissolve', 'fade', 'newsflash', 'plus',
        'pull', 'push', 'random', 'split', 'strips', 'wedge', 'wheel', 'wipe', 'zoom'].includes(type)) {
        transition.type = type;
      }
      // Parse direction from transition child attrs
      if (child.attrs['dir']) transition.direction = child.attrs['dir'];
    }
    if (child.tag === 'p:snd') {
      // Sound reference
      const sndRef = findChild(child, 'a:sndRef');
      if (sndRef && sndRef.attrs['r:id']) {
        transition.sound = sndRef.attrs['r:id'];
      }
    }
  }

  return transition;
}

function parseGroupShape(node: ParsedNode): GroupShape | null {
  const grpSpPr = findChild(node, 'p:grpSpPr');
  const position = grpSpPr ? parseGroupPosition(grpSpPr) : { x: 0, y: 0, width: 0, height: 0 };

  const children: SlideElement[] = [];
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'p:sp') {
      const shape = parseShape(child);
      if (shape) children.push(shape);
    } else if (child.tag === 'p:pic') {
      const media = parseMediaShape(child);
      if (media) {
        children.push(media);
      } else {
        const shape = parseImageShape(child);
        if (shape) children.push(shape);
      }
    } else if (child.tag === 'p:grpSp') {
      const group = parseGroupShape(child);
      if (group) children.push(group);
    } else if (child.tag === 'p:graphicFrame') {
      const table = parseTableShape(child);
      if (table) children.push(table);
    }
  }

  // Parse child offset and extent
  let childOffset: { x: number; y: number } | undefined;
  let childExtent: { width: number; height: number } | undefined;
  if (grpSpPr) {
    const xfrm = findChild(grpSpPr, 'a:xfrm');
    if (xfrm) {
      const chOff = findChild(xfrm, 'a:chOff');
      if (chOff) {
        childOffset = {
          x: chOff.attrs['x'] ? parseInt(chOff.attrs['x'], 10) : 0,
          y: chOff.attrs['y'] ? parseInt(chOff.attrs['y'], 10) : 0,
        };
      }
      const chExt = findChild(xfrm, 'a:chExt');
      if (chExt) {
        childExtent = {
          width: chExt.attrs['cx'] ? parseInt(chExt.attrs['cx'], 10) : 0,
          height: chExt.attrs['cy'] ? parseInt(chExt.attrs['cy'], 10) : 0,
        };
      }
    }
  }

  const result: GroupShape = { type: 'group', position, children };
  if (childOffset) result.childOffset = childOffset;
  if (childExtent) result.childExtent = childExtent;
  return result;
}

function parseMediaShape(node: ParsedNode): MediaShape | null {
  const nvPicPr = findChild(node, 'p:nvPicPr');
  if (!nvPicPr) return null;

  const nvPr = findChild(nvPicPr, 'p:nvPr');
  if (!nvPr) return null;

  const videoFile = findChild(nvPr, 'a:videoFile');
  const audioFile = findChild(nvPr, 'a:audioFile');
  const mediaFile = videoFile || audioFile;
  if (!mediaFile) return null;

  const relId = mediaFile.attrs['r:link'] || mediaFile.attrs['r:embed'] || '';
  const mediaType: 'video' | 'audio' = videoFile ? 'video' : 'audio';

  const spPr = findChild(node, 'p:spPr');
  const position = spPr ? parseMediaPosition(spPr) : { x: 0, y: 0, width: 0, height: 0 };

  return { type: 'media', mediaType, relationshipId: relId, position };
}

function parseMediaPosition(spPr: ParsedNode): Position {
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

function parseGroupPosition(grpSpPr: ParsedNode): Position {
  const xfrm = findChild(grpSpPr, 'a:xfrm');
  if (!xfrm) return { x: 0, y: 0, width: 0, height: 0 };

  const off = findChild(xfrm, 'a:off');
  const ext = findChild(xfrm, 'a:ext');

  const pos: Position = {
    x: off?.attrs['x'] ? parseInt(off.attrs['x'], 10) : 0,
    y: off?.attrs['y'] ? parseInt(off.attrs['y'], 10) : 0,
    width: ext?.attrs['cx'] ? parseInt(ext.attrs['cx'], 10) : 0,
    height: ext?.attrs['cy'] ? parseInt(ext.attrs['cy'], 10) : 0,
  };

  if (xfrm.attrs['rot']) {
    pos.rotation = parseInt(xfrm.attrs['rot'], 10) / 60000;
  }

  return pos;
}
