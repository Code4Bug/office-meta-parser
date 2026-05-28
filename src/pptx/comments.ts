import type { PptxPresentation, Slide, SlideComment } from './types.js';

export interface CommentInfo {
  comment: SlideComment;
  slideIndex: number;
}

/**
 * 生成下一个可用的批注 ID（从现有批注中推算最大 ID）
 */
function nextCommentId(pres: PptxPresentation): string {
  let maxId = 0;
  for (const slide of pres.slides) {
    if (slide.comments) {
      for (const c of slide.comments) {
        const num = parseInt(c.id, 10);
        if (!isNaN(num) && num > maxId) maxId = num;
      }
    }
  }
  return String(maxId + 1);
}

/**
 * 为指定幻灯片添加批注
 */
export function addComment(
  pres: PptxPresentation,
  slideIndex: number,
  authorName: string,
  text: string,
  x?: number,
  y?: number,
  date?: string,
): SlideComment {
  const slide = pres.slides[slideIndex];
  if (!slide) throw new Error(`Slide index ${slideIndex} out of range`);
  if (!slide.comments) slide.comments = [];

  const id = nextCommentId(pres);
  const comment: SlideComment = {
    id,
    authorId: 0,
    authorName,
    text,
    date: date || new Date().toISOString(),
  };
  if (x !== undefined && y !== undefined) {
    comment.position = { x, y };
  }

  slide.comments.push(comment);
  return comment;
}

/**
 * 移除指定幻灯片上的批注
 */
export function removeComment(pres: PptxPresentation, slideIndex: number, commentId: string): boolean {
  const slide = pres.slides[slideIndex];
  if (!slide?.comments) return false;

  const idx = slide.comments.findIndex(c => c.id === commentId);
  if (idx === -1) return false;

  slide.comments.splice(idx, 1);
  if (slide.comments.length === 0) delete slide.comments;
  return true;
}

/**
 * 列出演示文稿中所有批注
 */
export function listComments(pres: PptxPresentation): CommentInfo[] {
  const result: CommentInfo[] = [];
  for (let i = 0; i < pres.slides.length; i++) {
    const slide = pres.slides[i];
    if (!slide.comments) continue;
    for (const comment of slide.comments) {
      result.push({ comment, slideIndex: i });
    }
  }
  return result;
}

/**
 * 列出指定幻灯片的所有批注
 */
export function listSlideComments(pres: PptxPresentation, slideIndex: number): SlideComment[] {
  const slide = pres.slides[slideIndex];
  if (!slide?.comments) return [];
  return [...slide.comments];
}

/**
 * 获取指定批注的纯文本内容
 */
export function getCommentText(pres: PptxPresentation, slideIndex: number, commentId: string): string | null {
  const slide = pres.slides[slideIndex];
  if (!slide?.comments) return null;
  const comment = slide.comments.find(c => c.id === commentId);
  return comment?.text ?? null;
}
