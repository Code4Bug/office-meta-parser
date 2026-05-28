import type { ValidationIssue } from '../core/validate.js';
import type { DocxDocument, DocxBlock, Paragraph } from './types.js';

export function validateDocx(doc: DocxDocument): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  // ═══════════════ 必须项 (error) ═══════════════

  // 1. body.blocks 不能为空
  if (!doc.body?.blocks || doc.body.blocks.length === 0) {
    issues.push({ level: 'error', path: 'body.blocks', message: 'body.blocks cannot be empty' });
  }

  // 2. commentId 引用的 comment 必须存在
  const commentIds = new Set((doc.comments || []).map(c => c.id));
  forEachRun(doc, (run, path) => {
    if (run.commentId && !commentIds.has(run.commentId)) {
      issues.push({ level: 'error', path, message: `commentId "${run.commentId}" not found in comments` });
    }
  });

  // 3. bookmarkStart/bookmarkEnd id 成对
  const bmStartIds = new Set<string>();
  const bmEndIds = new Set<string>();
  forEachBlock(doc.body.blocks, 'body', (block) => {
    if (block.type === 'bookmarkStart') bmStartIds.add(block.id);
    if (block.type === 'bookmarkEnd') bmEndIds.add(block.id);
  });
  for (const id of bmStartIds) {
    if (!bmEndIds.has(id)) {
      issues.push({ level: 'error', path: 'body', message: `bookmarkStart id="${id}" has no matching bookmarkEnd` });
    }
  }
  for (const id of bmEndIds) {
    if (!bmStartIds.has(id)) {
      issues.push({ level: 'error', path: 'body', message: `bookmarkEnd id="${id}" has no matching bookmarkStart` });
    }
  }

  // 4. hyperlink 必须有 relationshipId
  forEachBlock(doc.body.blocks, 'body', (block, path) => {
    if (block.type === 'hyperlink' && !block.relationshipId) {
      issues.push({ level: 'error', path, message: 'hyperlink missing relationshipId' });
    }
  });

  // 5. image 必须有 relationshipId
  forEachBlock(doc.body.blocks, 'body', (block, path) => {
    if (block.type === 'image' && !block.relationshipId) {
      issues.push({ level: 'error', path, message: 'image missing relationshipId' });
    }
  });

  // 6. numbering numId 引用必须存在
  const numIds = new Set((doc.numbering?.nums || []).map(n => n.id));
  forEachParagraph(doc, (para, path) => {
    if (para.numbering && !numIds.has(para.numbering.numId)) {
      issues.push({ level: 'error', path, message: `numId "${para.numbering.numId}" not found in numbering definitions` });
    }
  });

  // 7. header/footer id 不能重复
  const headerIds = new Set<string>();
  for (const h of doc.headers || []) {
    if (headerIds.has(h.id)) {
      issues.push({ level: 'error', path: 'headers', message: `duplicate header id "${h.id}"` });
    }
    headerIds.add(h.id);
  }
  const footerIds = new Set<string>();
  for (const f of doc.footers || []) {
    if (footerIds.has(f.id)) {
      issues.push({ level: 'error', path: 'footers', message: `duplicate footer id "${f.id}"` });
    }
    footerIds.add(f.id);
  }

  // 8. table rows 不能为空
  forEachBlock(doc.body.blocks, 'body', (block, path) => {
    if (block.type === 'table' && (!block.rows || block.rows.length === 0)) {
      issues.push({ level: 'error', path, message: 'table has no rows' });
    }
  });

  // ═══════════════ 可选项 (warning) ═══════════════

  // 9. paragraph style 引用建议在 styles 中定义
  const styleIds = new Set((doc.styles?.paragraphStyles || []).map(s => s.id));
  forEachParagraph(doc, (para, path) => {
    if (para.style && !styleIds.has(para.style)) {
      issues.push({ level: 'warning', path, message: `paragraph style "${para.style}" not found in styleDefinitions` });
    }
  });

  // 10. comment content 不能为空
  for (const c of doc.comments || []) {
    if (!c.content || c.content.length === 0) {
      issues.push({ level: 'warning', path: `comments[${c.id}]`, message: `comment content is empty` });
    }
  }

  // 11. comments 有但正文没有 commentId 引用
  if ((doc.comments || []).length > 0) {
    let hasRef = false;
    forEachRun(doc, (run) => { if (run.commentId) hasRef = true; });
    if (!hasRef) {
      issues.push({ level: 'warning', path: 'body', message: 'comments defined but no commentId found in body runs' });
    }
  }

  return issues;
}

function forEachRun(doc: DocxDocument, fn: (run: import('./types.js').TextRun, path: string) => void) {
  const visit = (blocks: DocxBlock[], parentPath: string) => {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const path = `${parentPath}[${i}]`;
      if (block.type === 'paragraph') {
        for (let j = 0; j < block.runs.length; j++) {
          fn(block.runs[j], `${path}.runs[${j}]`);
        }
      } else if (block.type === 'table') {
        for (let ri = 0; ri < block.rows.length; ri++) {
          for (let ci = 0; ci < block.rows[ri].cells.length; ci++) {
            visit(block.rows[ri].cells[ci].blocks, `${path}.rows[${ri}].cells[${ci}]`);
          }
        }
      } else if (block.type === 'hyperlink') {
        for (let j = 0; j < block.runs.length; j++) {
          fn(block.runs[j], `${path}.runs[${j}]`);
        }
      }
    }
  };
  visit(doc.body.blocks, 'body.blocks');
}

function forEachBlock(blocks: DocxBlock[], parentPath: string, fn: (block: DocxBlock, path: string) => void) {
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const path = `${parentPath}[${i}]`;
    fn(block, path);
    if (block.type === 'table') {
      for (let ri = 0; ri < block.rows.length; ri++) {
        for (let ci = 0; ci < block.rows[ri].cells.length; ci++) {
          forEachBlock(block.rows[ri].cells[ci].blocks, `${path}.rows[${ri}].cells[${ci}]`, fn);
        }
      }
    }
  }
}

function forEachParagraph(doc: DocxDocument, fn: (para: Paragraph, path: string) => void) {
  const visit = (blocks: DocxBlock[], parentPath: string) => {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const path = `${parentPath}[${i}]`;
      if (block.type === 'paragraph') {
        fn(block, path);
      } else if (block.type === 'table') {
        for (let ri = 0; ri < block.rows.length; ri++) {
          for (let ci = 0; ci < block.rows[ri].cells.length; ci++) {
            visit(block.rows[ri].cells[ci].blocks, `${path}.rows[${ri}].cells[${ci}]`);
          }
        }
      }
    }
  };
  visit(doc.body.blocks, 'body.blocks');
}
