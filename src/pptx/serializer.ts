import type { PptxPresentation, SlideComment } from './types.js';
import {
  serializePresentation,
  serializePresentationRels,
  serializeSlideRels,
  serializeContentTypes,
} from './serializers/presentation.js';
import { serializeSlide } from './serializers/slide.js';
import { serializeSlideMaster, serializeSlideLayout, serializeMasterRels } from './serializers/master.js';
import { serializeTheme } from './serializers/theme.js';
import { serializeXml } from '../core/xml.js';
import type { ParsedNode } from '../core/types.js';

export interface PptxXmlParts {
  presentation: string;
  slides: string[];
  rels: string;
  slideRels: string[];
  contentTypes: string;
  masters: string[];
  masterRels: string[];
  layouts: string[];
  theme?: string;
  comments: { path: string; xml: string }[];
}

export function semanticToXml(pres: PptxPresentation): PptxXmlParts {
  const comments: { path: string; xml: string }[] = [];
  pres.slides.forEach((slide, i) => {
    if (slide.comments && slide.comments.length > 0) {
      comments.push({
        path: `ppt/comments/comment${i + 1}.xml`,
        xml: serializeCommentsXml(i, slide.comments),
      });
    }
  });

  return {
    presentation: serializePresentation(pres),
    slides: pres.slides.map((slide, i) => serializeSlide(slide, i)),
    rels: serializePresentationRels(pres.slides),
    slideRels: pres.slides.map((slide, i) => serializeSlideRels(i, !!slide.comments?.length)),
    contentTypes: serializeContentTypes(pres.slides, pres.layouts.length),
    masters: pres.masters.map((master, i) => serializeSlideMaster(master, i)),
    masterRels: pres.masters.map((master, i) => serializeMasterRels(master, i)),
    layouts: pres.layouts.map(layout => serializeSlideLayout(layout)),
    theme: pres.theme ? serializeTheme(pres.theme) : undefined,
    comments,
  };
}

function serializeCommentsXml(slideIndex: number, comments: SlideComment[]): string {
  const children: ParsedNode[] = [];

  // 收集所有作者
  const authors = new Set<string>();
  collectAuthors(comments, authors);
  const authorList = [...authors];

  // p:authorLst
  children.push({
    tag: 'p:authorLst',
    attrs: {},
    children: authorList.map(name => ({
      tag: 'p:author',
      attrs: { id: String(authorList.indexOf(name)), name },
      children: [],
    })),
  });

  // p:cmLst
  children.push({
    tag: 'p:cmLst',
    attrs: {},
    children: comments.map(c => serializeCommentNode(c, authorList)),
  });

  const root: ParsedNode = {
    tag: 'p:comments',
    attrs: {
      xmlns: 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeCommentNode(comment: SlideComment, authorList: string[]): ParsedNode {
  const attrs: Record<string, string> = {
    id: comment.id,
    authorId: String(authorList.indexOf(comment.authorName)),
    dt: comment.date || new Date().toISOString(),
  };
  if (comment.position) {
    attrs.x = String(comment.position.x);
    attrs.y = String(comment.position.y);
  }

  const children: ParsedNode[] = [
    { tag: 'p:text', attrs: {}, children: [comment.text] },
  ];

  // 嵌套回复
  if (comment.replies) {
    for (const reply of comment.replies) {
      children.push(serializeCommentNode(reply, authorList));
    }
  }

  return { tag: 'p:cm', attrs, children };
}

function collectAuthors(comments: SlideComment[], authors: Set<string>): void {
  for (const c of comments) {
    authors.add(c.authorName);
    if (c.replies) collectAuthors(c.replies, authors);
  }
}
