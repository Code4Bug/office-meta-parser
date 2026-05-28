import type { DocxDocument, DocxBlock, TextRun, Revision, RunProperties } from './types.js';

export interface RevisionInfo {
  revision: Revision;
  blockIndex: number;
  runIndex: number;
  run: TextRun;
}

/**
 * 标记 TextRun 为插入修订
 */
export function markInsert(run: TextRun, author: string, date?: string): void {
  run.revisionType = 'insert';
  run.revisionAuthor = author;
  run.revisionDate = date || new Date().toISOString();
}

/**
 * 标记 TextRun 为删除修订
 */
export function markDelete(run: TextRun, author: string, date?: string): void {
  run.revisionType = 'delete';
  run.revisionAuthor = author;
  run.revisionDate = date || new Date().toISOString();
}

/**
 * 添加格式变更修订记录到文档级别
 */
export function addFormatChange(doc: DocxDocument, author: string, date?: string): void {
  if (!doc.trackChanges) doc.trackChanges = [];
  doc.trackChanges.push({
    type: 'formatChange',
    author,
    date: date || new Date().toISOString(),
  });
}

/**
 * 清除 TextRun 上的修订标记
 */
export function clearRevision(run: TextRun): void {
  delete run.revisionType;
  delete run.revisionAuthor;
  delete run.revisionDate;
}

/**
 * 列出文档中所有修订及其位置
 */
export function listRevisions(doc: DocxDocument): RevisionInfo[] {
  const result: RevisionInfo[] = [];
  scanBlocks(doc.body.blocks, result);
  return result;
}

/**
 * 检查文档是否有未接受的修订
 */
export function hasPendingRevisions(doc: DocxDocument): boolean {
  return scanBlocksHasRevision(doc.body.blocks);
}

/**
 * 接受所有插入修订（移除 revisionType 标记，保留文本）
 */
export function acceptAllInserts(doc: DocxDocument): number {
  let count = 0;
  processBlocks(doc.body.blocks, 'insert', () => { count++; });
  return count;
}

/**
 * 接受所有删除修订（移除被标记为删除的 run）
 */
export function acceptAllDeletes(doc: DocxDocument): number {
  let count = 0;
  processBlocksDelete(doc.body.blocks, () => { count++; });
  return count;
}

/**
 * 拒绝所有插入修订（移除被标记为插入的 run）
 */
export function rejectAllInserts(doc: DocxDocument): number {
  let count = 0;
  processBlocksDelete(doc.body.blocks, () => { count++; }, 'insert');
  return count;
}

/**
 * 拒绝所有删除修订（移除 revisionType 标记，保留文本）
 */
export function rejectAllDeletes(doc: DocxDocument): number {
  let count = 0;
  processBlocks(doc.body.blocks, 'delete', () => { count++; });
  return count;
}

// ---- 内部辅助 ----

function scanBlocks(blocks: DocxBlock[], result: RevisionInfo[]): void {
  for (let bi = 0; bi < blocks.length; bi++) {
    const block = blocks[bi];
    if (block.type === 'paragraph') {
      for (let ri = 0; ri < block.runs.length; ri++) {
        const run = block.runs[ri];
        if (run.revisionType) {
          result.push({
            revision: {
              type: run.revisionType,
              author: run.revisionAuthor || '',
              date: run.revisionDate || '',
            },
            blockIndex: bi,
            runIndex: ri,
            run,
          });
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          scanBlocks(cell.blocks, result);
        }
      }
    }
  }
}

function scanBlocksHasRevision(blocks: DocxBlock[]): boolean {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      if (block.runs.some(r => r.revisionType)) return true;
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (scanBlocksHasRevision(cell.blocks)) return true;
        }
      }
    }
  }
  return false;
}

function processBlocks(blocks: DocxBlock[], type: 'insert' | 'delete', onProcess: () => void): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const run of block.runs) {
        if (run.revisionType === type) {
          onProcess();
          clearRevision(run);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          processBlocks(cell.blocks, type, onProcess);
        }
      }
    }
  }
}

function processBlocksDelete(blocks: DocxBlock[], onProcess: () => void, onlyType?: 'insert' | 'delete'): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      const original = block.runs.length;
      block.runs = block.runs.filter(run => {
        const shouldRemove = onlyType
          ? run.revisionType === onlyType
          : run.revisionType === 'delete';
        if (shouldRemove) onProcess();
        return !shouldRemove;
      });
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          processBlocksDelete(cell.blocks, onProcess, onlyType);
        }
      }
    }
  }
}
