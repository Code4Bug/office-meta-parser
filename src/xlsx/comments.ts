import type { XlsxWorkbook, Sheet, SheetComment, RichTextRun } from './types.js';

export interface CommentInfo {
  comment: SheetComment;
  sheetIndex: number;
}

/**
 * 为指定单元格添加批注
 */
export function addComment(
  wb: XlsxWorkbook,
  sheetIndex: number,
  ref: string,
  author: string,
  text: string,
  richText?: RichTextRun[],
): SheetComment {
  const sheet = wb.sheets[sheetIndex];
  if (!sheet) throw new Error(`Sheet index ${sheetIndex} out of range`);
  if (!sheet.comments) sheet.comments = [];

  // 如果已有同单元格批注，先移除
  const existing = sheet.comments.findIndex(c => c.ref === ref);
  if (existing !== -1) sheet.comments.splice(existing, 1);

  const authorId = ensureAuthor(wb, author);

  const comment: SheetComment = { ref, authorId, text };
  if (richText) comment.richText = richText;

  sheet.comments.push(comment);
  return comment;
}

/**
 * 移除指定单元格的批注
 */
export function removeComment(wb: XlsxWorkbook, sheetIndex: number, ref: string): boolean {
  const sheet = wb.sheets[sheetIndex];
  if (!sheet?.comments) return false;

  const idx = sheet.comments.findIndex(c => c.ref === ref);
  if (idx === -1) return false;

  sheet.comments.splice(idx, 1);
  return true;
}

/**
 * 列出工作簿中所有批注
 */
export function listComments(wb: XlsxWorkbook): CommentInfo[] {
  const result: CommentInfo[] = [];
  for (let i = 0; i < wb.sheets.length; i++) {
    const sheet = wb.sheets[i];
    if (!sheet.comments) continue;
    for (const comment of sheet.comments) {
      result.push({ comment, sheetIndex: i });
    }
  }
  return result;
}

/**
 * 列出指定 sheet 的所有批注
 */
export function listSheetComments(wb: XlsxWorkbook, sheetIndex: number): SheetComment[] {
  const sheet = wb.sheets[sheetIndex];
  if (!sheet?.comments) return [];
  return [...sheet.comments];
}

/**
 * 获取指定单元格的批注文本
 */
export function getCommentText(wb: XlsxWorkbook, sheetIndex: number, ref: string): string | null {
  const sheet = wb.sheets[sheetIndex];
  if (!sheet?.comments) return null;
  const comment = sheet.comments.find(c => c.ref === ref);
  return comment?.text ?? null;
}

/**
 * 更新指定单元格的批注内容
 */
export function updateComment(
  wb: XlsxWorkbook,
  sheetIndex: number,
  ref: string,
  text: string,
  richText?: RichTextRun[],
): boolean {
  const sheet = wb.sheets[sheetIndex];
  if (!sheet?.comments) return false;

  const comment = sheet.comments.find(c => c.ref === ref);
  if (!comment) return false;

  comment.text = text;
  if (richText !== undefined) comment.richText = richText;
  return true;
}

// ---- 内部辅助 ----

function ensureAuthor(wb: XlsxWorkbook, author: string): number {
  if (!wb.authors) wb.authors = [];
  const existing = wb.authors.indexOf(author);
  if (existing >= 0) return existing;
  wb.authors.push(author);
  return wb.authors.length - 1;
}
