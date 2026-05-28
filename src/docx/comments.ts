import type { DocxDocument, DocxBlock, Paragraph, TextRun, Comment, CommentExtended } from './types.js';

export interface CommentInfo {
  comment: Comment;
  blockIndex: number;
  runIndices: number[];
}

/**
 * 生成下一个可用的批注 ID（从现有批注中推算最大 ID）
 */
function nextCommentId(doc: DocxDocument): string {
  const existing = doc.comments || [];
  const maxId = existing.reduce((max, c) => {
    const num = parseInt(c.id, 10);
    return isNaN(num) ? max : Math.max(max, num);
  }, 0);
  return String(maxId + 1);
}

/**
 * 为指定 TextRun 添加批注
 * 会在 doc.comments 中创建批注记录，并在 run 上设置 commentId
 */
export function addComment(
  doc: DocxDocument,
  run: TextRun,
  author: string,
  content: string,
  date?: string,
): Comment {
  if (!doc.comments) doc.comments = [];

  const id = nextCommentId(doc);
  const comment: Comment = {
    id,
    author,
    date: date || new Date().toISOString(),
    content: [{ type: 'paragraph', runs: [{ text: content }] }],
  };

  doc.comments.push(comment);
  run.commentId = id;

  return comment;
}

/**
 * 移除指定批注
 * 同时清除所有引用该 commentId 的 run 标记
 */
export function removeComment(doc: DocxDocument, commentId: string): boolean {
  if (!doc.comments) return false;

  const idx = doc.comments.findIndex(c => c.id === commentId);
  if (idx === -1) return false;

  doc.comments.splice(idx, 1);

  // 清除 run 上的 commentId 标记
  clearCommentIdFromBlocks(doc.body.blocks, commentId);

  // 清除 commentExts
  if (doc.commentExts) {
    // commentExts 使用 paraId 关联，此处无法精确移除，保留
  }

  return true;
}

/**
 * 列出文档中所有批注及其位置信息
 */
export function listComments(doc: DocxDocument): CommentInfo[] {
  if (!doc.comments) return [];

  const result: CommentInfo[] = [];

  for (const comment of doc.comments) {
    const runIndices: number[] = [];
    findRunsWithCommentId(doc.body.blocks, comment.id, runIndices);
    // 取第一个 run 所在的 block index
    const blockIndex = findBlockIndexForComment(doc.body.blocks, comment.id);
    result.push({ comment, blockIndex, runIndices });
  }

  return result;
}

/**
 * 获取批注的纯文本内容
 */
export function getCommentText(doc: DocxDocument, commentId: string): string | null {
  const comment = doc.comments?.find(c => c.id === commentId);
  if (!comment) return null;
  return comment.content
    .flatMap(p => p.runs?.map(r => r.text) || [])
    .join('');
}

/**
 * 标记批注为已完成（设置 commentExts 的 done 标志）
 */
export function markCommentDone(doc: DocxDocument, commentId: string): boolean {
  // commentExts 按 paraId 索引，需要找到对应的 paraId
  // 在 OOXML 中 commentId 与 commentExtended 的 paraId 是同一个值
  if (!doc.commentExts) doc.commentExts = [];

  const ext = doc.commentExts.find(e => e.paraId === commentId);
  if (ext) {
    ext.done = true;
    return true;
  }

  // 如果不存在，创建一条
  doc.commentExts.push({ paraId: commentId, done: true });
  return true;
}

/**
 * 标记批注为未完成
 */
export function markCommentUndone(doc: DocxDocument, commentId: string): boolean {
  if (!doc.commentExts) return false;
  const ext = doc.commentExts.find(e => e.paraId === commentId);
  if (ext) {
    ext.done = false;
    return true;
  }
  return false;
}

// ---- 内部辅助 ----

function clearCommentIdFromBlocks(blocks: DocxBlock[], commentId: string): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const run of block.runs) {
        if (run.commentId === commentId) {
          delete run.commentId;
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          clearCommentIdFromBlocks(cell.blocks, commentId);
        }
      }
    }
  }
}

function findRunsWithCommentId(blocks: DocxBlock[], commentId: string, indices: number[]): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (let i = 0; i < block.runs.length; i++) {
        if (block.runs[i].commentId === commentId) {
          indices.push(i);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          findRunsWithCommentId(cell.blocks, commentId, indices);
        }
      }
    }
  }
}

function findBlockIndexForComment(blocks: DocxBlock[], commentId: string): number {
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      if (block.runs.some(r => r.commentId === commentId)) return i;
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          // 递归查找但返回当前 block index
          if (cell.blocks.some(b =>
            b.type === 'paragraph' && (b as Paragraph).runs?.some(r => r.commentId === commentId)
          )) return i;
        }
      }
    }
  }
  return -1;
}
