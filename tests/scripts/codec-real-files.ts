#!/usr/bin/env tsx
/**
 * 真实文件编解码测试脚本
 * 流程: 原始文件 → 解析 → 序列化 → 写入 tests/output/ → 读取 → 解析 → 序列化 → 再写入 → 再读取 → 解析
 * 对比三份语义模型 (original / file1 / file2) 的一致性
 */

import { mkdirSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { parseDocx, serializeDocx } from '../../src/docx/index.js';
import { parsePptx, serializePptx } from '../../src/pptx/index.js';
import { parseXlsx, serializeXlsx } from '../../src/xlsx/index.js';
import { loadFromFile, saveToFile } from '../../src/core/io.js';
import type { DocxDocument } from '../../src/docx/types.js';
import type { PptxPresentation } from '../../src/pptx/index.js';
import type { XlsxWorkbook } from '../../src/xlsx/types.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const OUTPUT_DIR = join(__dirname, '..', 'output');

// 确保输出目录存在
if (!existsSync(OUTPUT_DIR)) {
  mkdirSync(OUTPUT_DIR, { recursive: true });
}

// ============================================================
// 工具函数
// ============================================================

const INPUT_DIR = join(__dirname, '..', 'input');

function loadInputFile(name: string): Promise<ArrayBuffer> {
  return loadFromFile(join(INPUT_DIR, name));
}

function loadOutputFile(name: string): Promise<ArrayBuffer> {
  return loadFromFile(join(OUTPUT_DIR, name));
}

async function saveOutputFile(name: string, data: ArrayBuffer): Promise<void> {
  await saveToFile(data, join(OUTPUT_DIR, name));
}

function fmt(ms: number): string {
  return ms < 1 ? `${(ms * 1000).toFixed(0)}μs` : `${ms.toFixed(1)}ms`;
}

function sizeKb(bytes: number): string {
  return `${(bytes / 1024).toFixed(1)}KB`;
}

// ============================================================
// 语义模型打印
// ============================================================

function printDocxSemantic(label: string, doc: DocxDocument) {
  console.log(`\n  ┌─ ${label} ─────────────────────────────────`);

  const m = doc.meta;
  const metaFields = [
    m.title && `title="${m.title}"`,
    m.subject && `subject="${m.subject}"`,
    m.creator && `creator="${m.creator}"`,
    m.description && `desc="${m.description}"`,
    m.keywords && `keywords="${m.keywords}"`,
    m.lastModifiedBy && `lastModifiedBy="${m.lastModifiedBy}"`,
    m.category && `category="${m.category}"`,
  ].filter(Boolean);
  console.log(`  │ Meta: ${metaFields.length > 0 ? metaFields.join(', ') : '(空)'}`);

  if (doc.styles) {
    console.log(`  │ Styles: paragraph=${doc.styles.paragraphStyles?.length ?? 0}, character=${doc.styles.characterStyles?.length ?? 0}, table=${doc.styles.tableStyles?.length ?? 0}`);
  }

  if (doc.comments?.length) {
    console.log(`  │ Comments: ${doc.comments.length} 条`);
    for (const c of doc.comments) {
      const text = c.content?.flatMap(p => p.runs?.map(r => r.text) || []).join('') || '';
      console.log(`  │   [${c.id}] ${c.author} (${c.date}): "${text.slice(0, 80)}"`);
    }
  }

  if (doc.trackChanges?.length) {
    console.log(`  │ Revisions: ${doc.trackChanges.length} 条`);
    for (const r of doc.trackChanges) {
      console.log(`  │   ${r.type} by ${r.author} (${r.date})`);
    }
  }

  if (doc.numbering) {
    console.log(`  │ Numbering: abstractNums=${doc.numbering.abstractNums?.length ?? 0}, nums=${doc.numbering.nums?.length ?? 0}`);
  }

  if (doc.headers?.length) console.log(`  │ Headers: ${doc.headers.length} 个`);
  if (doc.footers?.length) console.log(`  │ Footers: ${doc.footers.length} 个`);
  if (doc.footnotes?.length) console.log(`  │ Footnotes: ${doc.footnotes.length} 条`);

  if (doc.body.sectionProperties) {
    const sp = doc.body.sectionProperties;
    const spFields = [
      sp.orientation && `orient=${sp.orientation}`,
      sp.pageWidth && `w=${sp.pageWidth}`,
      sp.pageHeight && `h=${sp.pageHeight}`,
      sp.pageNumberStart && `pgStart=${sp.pageNumberStart}`,
      sp.titlePage && `titlePage`,
      sp.verticalAlign && `vAlign=${sp.verticalAlign}`,
    ].filter(Boolean);
    console.log(`  │ Section: ${spFields.length > 0 ? spFields.join(', ') : '(有)'}`);
  }

  const blocks = doc.body.blocks;
  console.log(`  │ Body: ${blocks.length} 块`);
  const maxShow = 50;
  for (let i = 0; i < Math.min(blocks.length, maxShow); i++) {
    const b = blocks[i] as any;
    if (b.type === 'paragraph') {
      const text = (b.runs || []).map((r: any) => r.text).join('').slice(0, 60);
      const style = b.style ? ` [${b.style}]` : '';
      const num = b.numbering ? ` {num:${b.numbering.numId}/${b.numbering.level}}` : '';
      const props = [];
      if (b.properties?.alignment) props.push(`align=${b.properties.alignment}`);
      if (b.properties?.indent) props.push(`indent=${JSON.stringify(b.properties.indent)}`);
      if (b.properties?.spacing) props.push(`sp=${JSON.stringify(b.properties.spacing)}`);
      const propsStr = props.length > 0 ? ` <${props.join(', ')}>` : '';
      const runCount = b.runs?.length || 0;
      const runDetails = b.runs?.slice(0, 5).map((r: any) => {
        const flags = [
          r.bold && 'B', r.italic && 'I', r.underline && 'U', r.strike && 'S',
          r.superscript && '↑', r.subscript && '↓',
          r.fontSize && `${r.fontSize}pt`, r.color && `#${r.color}`,
          r.fontFamily && `${r.fontFamily}`,
          r.highlight && `hl=${r.highlight}`,
        ].filter(Boolean).join(',');
        return flags ? `{${flags}}` : '';
      }).filter(Boolean);
      const runStr = runDetails?.length ? ` [${runDetails.join(' ')}]` : '';
      console.log(`  │ [${i}] paragraph${style}${num}${propsStr}: "${text}"${runStr} (${runCount} runs)`);
    } else if (b.type === 'table') {
      const rows = b.rows?.length || 0;
      const cols = b.rows?.[0]?.cells?.length || 0;
      const tblStyle = b.properties?.style ? ` [${b.properties.style}]` : '';
      console.log(`  │ [${i}] table${tblStyle}: ${rows}×${cols}`);
      for (let r = 0; r < Math.min(rows, 3); r++) {
        const cells = b.rows[r]?.cells || [];
        const cellTexts = cells.slice(0, 6).map((c: any) => {
          const t = c.blocks?.[0]?.runs?.map((r: any) => r.text).join('') || '';
          return `"${t.slice(0, 20)}"`;
        });
        const suffix = cells.length > 6 ? ` ...+${cells.length - 6}` : '';
        console.log(`  │   row[${r}]: ${cellTexts.join(' | ')}${suffix}`);
      }
    } else if (b.type === 'image') {
      console.log(`  │ [${i}] image: ${b.width}×${b.height} rel=${b.relationshipId}${b.alt ? ` alt="${b.alt}"` : ''}${b.isFloating ? ' floating' : ''}`);
    } else if (b.type === 'hyperlink') {
      const text = b.runs?.[0]?.text || '';
      console.log(`  │ [${i}] hyperlink: "${text}" rel=${b.relationshipId}${b.tooltip ? ` tip="${b.tooltip}"` : ''}`);
    } else if (b.type === 'bookmarkStart') {
      console.log(`  │ [${i}] bookmarkStart: "${b.name}" id=${b.id}`);
    } else if (b.type === 'bookmarkEnd') {
      console.log(`  │ [${i}] bookmarkEnd: id=${b.id}`);
    } else {
      console.log(`  │ [${i}] ${b.type}`);
    }
  }
  if (blocks.length > maxShow) {
    console.log(`  │ ... 还有 ${blocks.length - maxShow} 块`);
  }

  console.log(`  └──────────────────────────────────────────────`);
}

function printXlsxSemantic(label: string, wb: XlsxWorkbook) {
  console.log(`\n  ┌─ ${label} ─────────────────────────────────`);

  const m = wb.meta;
  const metaFields = [
    m.title && `title="${m.title}"`,
    m.creator && `creator="${m.creator}"`,
  ].filter(Boolean);
  console.log(`  │ Meta: ${metaFields.length > 0 ? metaFields.join(', ') : '(空)'}`);
  console.log(`  │ SharedStrings: ${wb.sharedStrings.length} 个`);

  for (let i = 0; i < wb.sheets.length; i++) {
    const s = wb.sheets[i];
    console.log(`  │ Sheet[${i}] "${s.name}" state=${s.state ?? 'visible'}: ${s.cells.length} 行`);

    if (s.mergedCells.length > 0) {
      console.log(`  │   mergedCells: ${s.mergedCells.length} 个`);
    }
    if (s.columnWidths.length > 0) {
      console.log(`  │   columnWidths: [${s.columnWidths.slice(0, 8).join(', ')}${s.columnWidths.length > 8 ? '...' : ''}]`);
    }
    if (s.rowHeights.length > 0) {
      console.log(`  │   rowHeights: ${s.rowHeights.length} 行自定义`);
    }
    if (s.frozenPanes) {
      console.log(`  │   frozenPanes: xSplit=${s.frozenPanes.xSplit}, ySplit=${s.frozenPanes.ySplit}`);
    }
    if (s.printArea) {
      const pa = s.printArea;
      const paFields = [
        pa.fitToWidth && `fitW=${pa.fitToWidth}`,
        pa.fitToHeight && `fitH=${pa.fitToHeight}`,
        pa.pageMargins && `margins=${pa.pageMargins.top}/${pa.pageMargins.right}/${pa.pageMargins.bottom}/${pa.pageMargins.left}`,
      ].filter(Boolean);
      console.log(`  │   printArea: ${paFields.join(', ') || '(有)'}`);
    }
    if (s.conditionalFormats?.length) console.log(`  │   conditionalFormats: ${s.conditionalFormats.length} 个`);
    if (s.dataValidations?.length) console.log(`  │   dataValidations: ${s.dataValidations.length} 个`);
    if (s.autoFilter) console.log(`  │   autoFilter: ${s.autoFilter.ref}`);
    if (s.protection) console.log(`  │   protection: ${JSON.stringify(s.protection)}`);

    const maxRows = Math.min(s.cells.length, 10);
    for (let r = 0; r < maxRows; r++) {
      const row = s.cells[r];
      if (!row) continue;
      const maxCols = Math.min(row.length, 10);
      const cells: string[] = [];
      for (let c = 0; c < maxCols; c++) {
        const cell = row[c];
        if (!cell) continue;
        let v = cell.value !== undefined ? String(cell.value) : '';
        if (v.length > 15) v = v.slice(0, 15) + '…';
        if (cell.formula) v = `=${cell.formula}`;
        if (cell.type === 'boolean') v = cell.value ? 'TRUE' : 'FALSE';
        if (cell.type === 'error') v = `#${cell.value}`;
        cells.push(`${getColLetter(c)}${r + 1}:${v}`);
      }
      if (cells.length > 0) {
        const suffix = row.length > maxCols ? ` ...+${row.length - maxCols}cols` : '';
        console.log(`  │   ${cells.join(' | ')}${suffix}`);
      }
    }
    if (s.cells.length > maxRows) {
      console.log(`  │   ... 还有 ${s.cells.length - maxRows} 行`);
    }
  }

  console.log(`  └──────────────────────────────────────────────`);
}

function printPptxSemantic(label: string, pres: PptxPresentation) {
  console.log(`\n  ┌─ ${label} ─────────────────────────────────`);

  const m = pres.meta;
  const metaFields = [
    m.title && `title="${m.title}"`,
    m.subject && `subject="${m.subject}"`,
    m.creator && `creator="${m.creator}"`,
    m.description && `desc="${m.description}"`,
  ].filter(Boolean);
  console.log(`  │ Meta: ${metaFields.length > 0 ? metaFields.join(', ') : '(空)'}`);

  if (pres.slideSize) {
    console.log(`  │ SlideSize: ${pres.slideSize.width}×${pres.slideSize.height}`);
  }
  console.log(`  │ Masters: ${pres.masters.length}, Layouts: ${pres.layouts.length}`);
  console.log(`  │ Slides: ${pres.slides.length}`);

  for (let i = 0; i < pres.slides.length; i++) {
    const s = pres.slides[i];
    console.log(`  │ Slide[${i}]: ${s.elements.length} 元素${s.transition ? `, transition=${s.transition.type}` : ''}${s.notes ? ', 有备注' : ''}`);

    for (let j = 0; j < Math.min(s.elements.length, 10); j++) {
      const el = s.elements[j];
      if (el.type === 'text') {
        const text = el.content?.slice(0, 50) || '';
        const pos = el.position ? ` @(${el.position.x},${el.position.y} ${el.position.width}×${el.position.height})` : '';
        const ph = el.placeholder ? ` [${el.placeholder.type}${el.placeholder.index !== undefined ? `#${el.placeholder.index}` : ''}]` : '';
        console.log(`  │   [${j}] text${ph}${pos}: "${text}"`);
      } else if (el.type === 'image') {
        const pos = el.position ? ` @(${el.position.x},${el.position.y} ${el.position.width}×${el.position.height})` : '';
        console.log(`  │   [${j}] image${pos} rel=${el.relationshipId}${el.alt ? ` alt="${el.alt}"` : ''}`);
      } else if (el.type === 'group') {
        console.log(`  │   [${j}] group: ${el.children?.length ?? 0} 子元素`);
      } else if (el.type === 'table') {
        console.log(`  │   [${j}] table: ${el.rows?.length ?? 0} 行`);
      } else {
        console.log(`  │   [${j}] ${el.type}`);
      }
    }
    if (s.elements.length > 10) {
      console.log(`  │   ... 还有 ${s.elements.length - 10} 元素`);
    }
  }

  console.log(`  └──────────────────────────────────────────────`);
}

function getColLetter(col: number): string {
  let s = '';
  let c = col;
  while (c >= 0) {
    s = String.fromCharCode(65 + (c % 26)) + s;
    c = Math.floor(c / 26) - 1;
  }
  return s;
}

// ============================================================
// 属性对比工具
// ============================================================

interface CheckResult {
  category: string;
  item: string;
  orig: string;
  restored: string;
  match: boolean;
}

function toStr(v: any): string {
  return v === undefined ? 'undefined' : typeof v === 'object' ? JSON.stringify(v) : String(v);
}

function createChecker(details: CheckResult[]) {
  return (category: string, item: string, orig: any, rest: any) => {
    const match = toStr(orig) === toStr(rest);
    details.push({ category, item, orig: toStr(orig), restored: toStr(rest), match });
  };
}

function compareDocxSemantic(
  a: DocxDocument, b: DocxDocument, details: CheckResult[],
) {
  const check = createChecker(details);

  check('Meta', 'title', a.meta.title, b.meta.title);
  check('Meta', 'subject', a.meta.subject, b.meta.subject);
  check('Meta', 'creator', a.meta.creator, b.meta.creator);
  check('Meta', 'description', a.meta.description, b.meta.description);
  check('Meta', 'keywords', a.meta.keywords, b.meta.keywords);
  check('Meta', 'lastModifiedBy', a.meta.lastModifiedBy, b.meta.lastModifiedBy);
  check('Meta', 'category', a.meta.category, b.meta.category);

  check('Body', 'blocks.length', a.body.blocks.length, b.body.blocks.length);

  const aBlocks = a.body.blocks;
  const bBlocks = b.body.blocks;
  const maxBlocks = Math.min(aBlocks.length, bBlocks.length, 200);

  for (let i = 0; i < maxBlocks; i++) {
    const ab = aBlocks[i] as any;
    const bb = bBlocks[i] as any;
    check(`Block[${i}]`, 'type', ab.type, bb.type);

    if (ab.type === 'paragraph' && bb.type === 'paragraph') {
      if (ab.properties || bb.properties) {
        check(`Block[${i}].para`, 'alignment', ab.properties?.alignment, bb.properties?.alignment);
        check(`Block[${i}].para`, 'indent', ab.properties?.indent, bb.properties?.indent);
        check(`Block[${i}].para`, 'spacing', ab.properties?.spacing, bb.properties?.spacing);
        check(`Block[${i}].para`, 'keepNext', ab.properties?.keepNext, bb.properties?.keepNext);
        check(`Block[${i}].para`, 'keepLines', ab.properties?.keepLines, bb.properties?.keepLines);
        check(`Block[${i}].para`, 'pageBreakBefore', ab.properties?.pageBreakBefore, bb.properties?.pageBreakBefore);
        check(`Block[${i}].para`, 'outlineLevel', ab.properties?.outlineLevel, bb.properties?.outlineLevel);
        check(`Block[${i}].para`, 'border', ab.properties?.border, bb.properties?.border);
        check(`Block[${i}].para`, 'shading', ab.properties?.shading, bb.properties?.shading);
      }
      check(`Block[${i}]`, 'style', ab.style, bb.style);

      if (ab.numbering || bb.numbering) {
        check(`Block[${i}].num`, 'level', ab.numbering?.level, bb.numbering?.level);
        check(`Block[${i}].num`, 'numId', ab.numbering?.numId, bb.numbering?.numId);
      }

      check(`Block[${i}]`, 'runs.length', ab.runs?.length, bb.runs?.length);
      const runCount = Math.min(ab.runs?.length || 0, bb.runs?.length || 0, 50);
      for (let j = 0; j < runCount; j++) {
        const ar = ab.runs[j];
        const br = bb.runs[j];
        const pfx = `Block[${i}].run[${j}]`;
        check(pfx, 'text', ar.text, br.text);
        check(pfx, 'bold', ar.bold, br.bold);
        check(pfx, 'italic', ar.italic, br.italic);
        check(pfx, 'underline', ar.underline, br.underline);
        check(pfx, 'strike', ar.strike, br.strike);
        check(pfx, 'superscript', ar.superscript, br.superscript);
        check(pfx, 'subscript', ar.subscript, br.subscript);
        check(pfx, 'fontSize', ar.fontSize, br.fontSize);
        check(pfx, 'color', ar.color, br.color);
        check(pfx, 'fontFamily', ar.fontFamily, br.fontFamily);
        check(pfx, 'highlight', ar.highlight, br.highlight);
        check(pfx, 'caps', ar.caps, br.caps);
        check(pfx, 'smallCaps', ar.smallCaps, br.smallCaps);
        check(pfx, 'dstrike', ar.dstrike, br.dstrike);
        check(pfx, 'vanish', ar.vanish, br.vanish);
        check(pfx, 'characterSpacing', ar.characterSpacing, br.characterSpacing);
        check(pfx, 'emphasis', ar.emphasis, br.emphasis);
        check(pfx, 'verticalPosition', ar.verticalPosition, br.verticalPosition);
      }
    } else if (ab.type === 'table' && bb.type === 'table') {
      check(`Block[${i}]`, 'rows.length', ab.rows?.length, bb.rows?.length);
      if (ab.properties || bb.properties) {
        check(`Block[${i}].tbl`, 'style', ab.properties?.style, bb.properties?.style);
        check(`Block[${i}].tbl`, 'width', ab.properties?.width, bb.properties?.width);
        check(`Block[${i}].tbl`, 'borders', ab.properties?.borders, bb.properties?.borders);
        check(`Block[${i}].tbl`, 'layout', ab.properties?.layout, bb.properties?.layout);
      }
      const rowCount = Math.min(ab.rows?.length || 0, bb.rows?.length || 0, 3);
      for (let r = 0; r < rowCount; r++) {
        const acells = ab.rows[r]?.cells || [];
        const bcells = bb.rows[r]?.cells || [];
        check(`Block[${i}].row[${r}]`, 'cells.length', acells.length, bcells.length);
        for (let c = 0; c < Math.min(acells.length, bcells.length, 5); c++) {
          const at = acells[c]?.blocks?.[0]?.runs?.map((r: any) => r.text).join('');
          const bt = bcells[c]?.blocks?.[0]?.runs?.map((r: any) => r.text).join('');
          check(`Block[${i}].row[${r}].cell[${c}]`, 'text', at, bt);
        }
      }
    } else if (ab.type === 'image' && bb.type === 'image') {
      check(`Block[${i}]`, 'relationshipId', ab.relationshipId, bb.relationshipId);
      check(`Block[${i}]`, 'width', ab.width, bb.width);
      check(`Block[${i}]`, 'height', ab.height, bb.height);
      check(`Block[${i}]`, 'alt', ab.alt, bb.alt);
      check(`Block[${i}]`, 'isFloating', ab.isFloating, bb.isFloating);
    } else if (ab.type === 'hyperlink' && bb.type === 'hyperlink') {
      check(`Block[${i}]`, 'relationshipId', ab.relationshipId, bb.relationshipId);
      check(`Block[${i}]`, 'tooltip', ab.tooltip, bb.tooltip);
      check(`Block[${i}]`, 'runs[0].text', ab.runs?.[0]?.text, bb.runs?.[0]?.text);
    }
  }

  if (a.styles && b.styles) {
    check('Styles', 'paragraphStyles', a.styles.paragraphStyles?.length, b.styles.paragraphStyles?.length);
    check('Styles', 'characterStyles', a.styles.characterStyles?.length, b.styles.characterStyles?.length);
    check('Styles', 'tableStyles', a.styles.tableStyles?.length, b.styles.tableStyles?.length);
  }

  if (a.comments || b.comments) {
    check('Comments', 'count', a.comments?.length, b.comments?.length);
    const cCount = Math.min(a.comments?.length || 0, b.comments?.length || 0);
    for (let ci = 0; ci < cCount; ci++) {
      const ac = a.comments![ci];
      const bc = b.comments![ci];
      const cp = `Comment[${ci}]`;
      check(cp, 'author', ac.author, bc.author);
      check(cp, 'date', ac.date, bc.date);
      const aText = ac.content?.flatMap(p => p.runs?.map(r => r.text) || []).join('') || '';
      const bText = bc.content?.flatMap(p => p.runs?.map(r => r.text) || []).join('') || '';
      check(cp, 'text', aText, bText);
    }
  }

  if (a.trackChanges || b.trackChanges) {
    check('Revisions', 'count', a.trackChanges?.length, b.trackChanges?.length);
    const rCount = Math.min(a.trackChanges?.length || 0, b.trackChanges?.length || 0);
    for (let ri = 0; ri < rCount; ri++) {
      const ar = a.trackChanges![ri];
      const br = b.trackChanges![ri];
      check(`Rev[${ri}]`, 'type', ar.type, br.type);
      check(`Rev[${ri}]`, 'author', ar.author, br.author);
      check(`Rev[${ri}]`, 'date', ar.date, br.date);
    }
  }

  if (a.headers || b.headers) {
    check('Headers', 'count', a.headers?.length, b.headers?.length);
  }
  if (a.numbering || b.numbering) {
    check('Numbering', 'abstractNums', a.numbering?.abstractNums?.length, b.numbering?.abstractNums?.length);
    check('Numbering', 'nums', a.numbering?.nums?.length, b.numbering?.nums?.length);
  }
  if (a.footnotes || b.footnotes) {
    check('Footnotes', 'count', a.footnotes?.length, b.footnotes?.length);
  }
  if (a.body.sectionProperties || b.body.sectionProperties) {
    const sp = a.body.sectionProperties;
    const rp = b.body.sectionProperties;
    check('Section', 'orientation', sp?.orientation, rp?.orientation);
    check('Section', 'pageNumberStart', sp?.pageNumberStart, rp?.pageNumberStart);
    check('Section', 'titlePage', sp?.titlePage, rp?.titlePage);
    check('Section', 'verticalAlign', sp?.verticalAlign, rp?.verticalAlign);
  }
}

function compareXlsxSemantic(
  a: XlsxWorkbook, b: XlsxWorkbook, details: CheckResult[],
) {
  const check = createChecker(details);

  check('Meta', 'title', a.meta.title, b.meta.title);
  check('Meta', 'creator', a.meta.creator, b.meta.creator);

  check('Workbook', 'sheets.length', a.sheets.length, b.sheets.length);
  check('Workbook', 'sharedStrings.length', a.sharedStrings.length, b.sharedStrings.length);

  for (let i = 0; i < Math.min(a.sheets.length, b.sheets.length); i++) {
    const os = a.sheets[i];
    const rs = b.sheets[i];
    const p = `Sheet[${i}]`;

    check(p, 'name', os.name, rs.name);
    check(p, 'state', os.state, rs.state);
    check(p, 'tabColor', os.tabColor, rs.tabColor);
    check(p, 'cells.length', os.cells.length, rs.cells.length);
    check(p, 'mergedCells.length', os.mergedCells.length, rs.mergedCells.length);
    check(p, 'columnWidths.length', os.columnWidths.length, rs.columnWidths.length);
    check(p, 'rowHeights.length', os.rowHeights.length, rs.rowHeights.length);
    check(p, 'hyperlinks.length', os.hyperlinks.length, rs.hyperlinks.length);

    check(p, 'frozenPanes', os.frozenPanes, rs.frozenPanes);
    check(p, 'zoomScale', os.zoomScale, rs.zoomScale);
    check(p, 'activeCell', os.activeCell, rs.activeCell);

    check(p, 'defaultRowHeight', os.defaultRowHeight, rs.defaultRowHeight);
    check(p, 'defaultColWidth', os.defaultColWidth, rs.defaultColWidth);

    check(p, 'rowGroups', os.rowGroups, rs.rowGroups);
    check(p, 'colGroups', os.colGroups, rs.colGroups);

    check(p, 'pageSetup', os.pageSetup, rs.pageSetup);
    check(p, 'headerFooter', os.headerFooter, rs.headerFooter);

    check(p, 'conditionalFormats', os.conditionalFormats, rs.conditionalFormats);
    check(p, 'dataValidations', os.dataValidations, rs.dataValidations);
    check(p, 'autoFilter', os.autoFilter, rs.autoFilter);

    check(p, 'protection', os.protection, rs.protection);
    check(p, 'printArea', os.printArea, rs.printArea);
    check(p, 'tables', os.tables, rs.tables);

    const maxRows = Math.min(os.cells.length, rs.cells.length, 10);
    for (let r = 0; r < maxRows; r++) {
      const maxCols = Math.min(os.cells[r]?.length || 0, rs.cells[r]?.length || 0, 20);
      for (let c = 0; c < maxCols; c++) {
        const oc = os.cells[r]?.[c];
        const rc = rs.cells[r]?.[c];
        if (!oc && !rc) continue;
        check(`${p}.cell[${r},${c}]`, 'value', oc?.value, rc?.value);
        check(`${p}.cell[${r},${c}]`, 'type', oc?.type, rc?.type);
        if (oc?.formula) check(`${p}.cell[${r},${c}]`, 'formula', oc.formula, rc?.formula);
      }
    }
  }
}

function comparePptxSemantic(
  a: PptxPresentation, b: PptxPresentation, details: CheckResult[],
) {
  const check = createChecker(details);

  check('Meta', 'title', a.meta.title, b.meta.title);
  check('Meta', 'subject', a.meta.subject, b.meta.subject);
  check('Meta', 'creator', a.meta.creator, b.meta.creator);
  check('Meta', 'description', a.meta.description, b.meta.description);

  check('Presentation', 'slides.length', a.slides.length, b.slides.length);
  check('Presentation', 'masters.length', a.masters.length, b.masters.length);
  check('Presentation', 'layouts.length', a.layouts.length, b.layouts.length);

  const maxSlides = Math.min(a.slides.length, b.slides.length);
  for (let i = 0; i < maxSlides; i++) {
    const as = a.slides[i];
    const bs = b.slides[i];
    const sp = `Slide[${i}]`;

    check(sp, 'elements.length', as.elements.length, bs.elements.length);
    check(sp, 'transition', as.transition?.type, bs.transition?.type);

    const maxEls = Math.min(as.elements.length, bs.elements.length, 20);
    for (let j = 0; j < maxEls; j++) {
      const ae = as.elements[j];
      const be = bs.elements[j];
      const ep = `${sp}.el[${j}]`;

      check(ep, 'type', ae.type, be.type);
      if (ae.type === 'text' && be.type === 'text') {
        check(ep, 'content', ae.content?.slice(0, 200), be.content?.slice(0, 200));
        check(ep, 'position', ae.position, be.position);
        check(ep, 'placeholder', ae.placeholder?.type, be.placeholder?.type);
      } else if (ae.type === 'image' && be.type === 'image') {
        check(ep, 'relationshipId', ae.relationshipId, be.relationshipId);
        check(ep, 'position', ae.position, be.position);
        check(ep, 'alt', ae.alt, be.alt);
      }
    }
  }
}

// ============================================================
// Raw 解码数据对比
// ============================================================

interface RawDiffResult {
  part: string;
  status: 'same' | 'changed' | 'added' | 'removed';
  details?: string;
}

function compareParsedNodes(a: ParsedNode, b: ParsedNode, path: string, maxDiffs = 50): string[] {
  const diffs: string[] = [];

  if (a.tag !== b.tag) {
    diffs.push(`${path}: tag "${a.tag}" → "${b.tag}"`);
    return diffs;
  }

  // 对比属性 (忽略顺序)
  const aKeys = Object.keys(a.attrs).sort();
  const bKeys = Object.keys(b.attrs).sort();
  const allKeys = new Set([...aKeys, ...bKeys]);
  for (const key of allKeys) {
    const av = a.attrs[key];
    const bv = b.attrs[key];
    if (av !== bv) {
      diffs.push(`${path}@${key}: "${av ?? '(无)'}" → "${bv ?? '(无)'}"`);
    }
  }

  // 对比子节点
  const maxLen = Math.max(a.children.length, b.children.length);
  for (let i = 0; i < maxLen; i++) {
    const ac = a.children[i];
    const bc = b.children[i];
    const childPath = `${path}/${i}`;

    if (ac === undefined) {
      diffs.push(`${childPath}: 缺失 (原为 ${typeof bc === 'string' ? `"${truncate(bc, 30)}"` : bc?.tag})`);
    } else if (bc === undefined) {
      diffs.push(`${childPath}: 多余 (值为 ${typeof ac === 'string' ? `"${truncate(ac, 30)}"` : ac?.tag})`);
    } else if (typeof ac === 'string' && typeof bc === 'string') {
      if (ac !== bc) {
        diffs.push(`${childPath}: 文本 "${truncate(ac, 30)}" → "${truncate(bc, 30)}"`);
      }
    } else if (typeof ac === 'string' || typeof bc === 'string') {
      diffs.push(`${childPath}: 类型不同 (${typeof ac === 'string' ? '文本' : ac?.tag} → ${typeof bc === 'string' ? '文本' : bc?.tag})`);
    } else {
      diffs.push(...compareParsedNodes(ac, bc, childPath, maxDiffs - diffs.length));
    }

    if (diffs.length >= maxDiffs) break;
  }

  return diffs;
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 3) + '...' : s;
}

function compareRawDocuments(a: RawDocument, b: RawDocument): RawDiffResult[] {
  const results: RawDiffResult[] = [];

  // 对比 contentTypes
  const aCt = JSON.stringify(a.contentTypes.sort((x, y) => x.partName.localeCompare(y.partName)));
  const bCt = JSON.stringify(b.contentTypes.sort((x, y) => x.partName.localeCompare(y.partName)));
  if (aCt !== bCt) {
    results.push({ part: '[Content_Types].xml', status: 'changed', details: '内容类型不同' });
  } else {
    results.push({ part: '[Content_Types].xml', status: 'same' });
  }

  // 对比 rels
  const aRelsKeys = [...a.rels.keys()].sort();
  const bRelsKeys = [...b.rels.keys()].sort();
  const allRelsKeys = new Set([...aRelsKeys, ...bRelsKeys]);
  for (const key of allRelsKeys) {
    const ar = a.rels.get(key);
    const br = b.rels.get(key);
    if (!ar) {
      results.push({ part: key, status: 'added', details: '输出文件新增 rels' });
    } else if (!br) {
      results.push({ part: key, status: 'removed', details: '输出文件缺少 rels' });
    } else {
      const aJson = JSON.stringify(ar.sort((x, y) => x.id.localeCompare(y.id)));
      const bJson = JSON.stringify(br.sort((x, y) => x.id.localeCompare(y.id)));
      if (aJson !== bJson) {
        results.push({ part: key, status: 'changed', details: '关系引用不同' });
      } else {
        results.push({ part: key, status: 'same' });
      }
    }
  }

  // 对比 XML parts
  const aPartsKeys = [...a.parts.keys()].sort();
  const bPartsKeys = [...b.parts.keys()].sort();
  const allPartsKeys = new Set([...aPartsKeys, ...bPartsKeys]);

  for (const key of allPartsKeys) {
    const ap = a.parts.get(key);
    const bp = b.parts.get(key);

    if (!ap) {
      results.push({ part: key, status: 'added', details: '输出文件新增 XML part' });
    } else if (!bp) {
      results.push({ part: key, status: 'removed', details: '输出文件缺少 XML part' });
    } else {
      const diffs = compareParsedNodes(ap, bp, key);
      if (diffs.length === 0) {
        results.push({ part: key, status: 'same' });
      } else {
        // 对关键 part 保存具体差异
        const isKey = key.includes('document.xml') || key.includes('styles.xml') || key.includes('sharedStrings');
        const detailText = isKey
          ? `${diffs.length} 处差异 (如: ${diffs.slice(0, 3).map(d => truncate(d, 50)).join('; ')})`
          : `${diffs.length} 处差异`;
        results.push({ part: key, status: 'changed', details: detailText });
      }
    }
  }

  return results;
}

// ============================================================
// 报告结构
// ============================================================

interface Report {
  file: string;
  fileSize: string;
  outputSize1: string;
  outputSize2: string;
  timings: { parse1: string; serialize1: string; parse2: string; serialize2: string; parse3: string };
  roundTrip1: CheckResult[];  // original vs file1
  roundTrip2: CheckResult[];  // file1 vs file2
  rawDiff1: RawDiffResult[];  // raw 原始 vs raw 文件1
  rawDiff2: RawDiffResult[];  // raw 文件1 vs raw 文件2
  original?: DocxDocument | PptxPresentation | XlsxWorkbook;
  file1?: DocxDocument | PptxPresentation | XlsxWorkbook;
  file2?: DocxDocument | PptxPresentation | XlsxWorkbook;
  rawOriginal?: RawDocument;
  rawFile1?: RawDocument;
  rawFile2?: RawDocument;
}

// ============================================================
// DOCX 测试
// ============================================================

async function testDocxFile(fileName: string): Promise<Report> {
  const report: Report = {
    file: fileName,
    fileSize: '',
    outputSize1: '',
    outputSize2: '',
    timings: { parse1: '', serialize1: '', parse2: '', serialize2: '', parse3: '' },
    roundTrip1: [],
    roundTrip2: [],
    rawDiff1: [],
    rawDiff2: [],
  };

  try {
    const buffer = await loadInputFile(fileName);
    report.fileSize = sizeKb(buffer.byteLength);

    // 第 1 步: 原始文件 → 解析
    const t0 = performance.now();
    const { raw: rawOriginal, semantic: original } = await parseDocx(buffer);
    const t1 = performance.now();
    report.timings.parse1 = fmt(t1 - t0);
    report.original = original;
    report.rawOriginal = rawOriginal;

    // 第 2 步: 序列化 → 写入文件
    const t2 = performance.now();
    const output1 = await serializeDocx(original);
    const t3 = performance.now();
    report.timings.serialize1 = fmt(t3 - t2);
    report.outputSize1 = sizeKb(output1.byteLength);
    await saveOutputFile(fileName, output1);

    // 第 3 步: 读取文件 → 解析
    const t4 = performance.now();
    const { raw: rawFile1, semantic: file1 } = await parseDocx(await loadOutputFile(fileName));
    const t5 = performance.now();
    report.timings.parse2 = fmt(t5 - t4);
    report.file1 = file1;
    report.rawFile1 = rawFile1;

    // 第 4 步: 再次序列化 → 写入文件
    const t6 = performance.now();
    const output2 = await serializeDocx(file1);
    const t7 = performance.now();
    report.timings.serialize2 = fmt(t7 - t6);
    report.outputSize2 = sizeKb(output2.byteLength);
    await saveOutputFile(fileName, output2);

    // 第 5 步: 再次读取文件 → 解析
    const t8 = performance.now();
    const { raw: rawFile2, semantic: file2 } = await parseDocx(await loadOutputFile(fileName));
    const t9 = performance.now();
    report.timings.parse3 = fmt(t9 - t8);
    report.file2 = file2;
    report.rawFile2 = rawFile2;

    // 语义对比: original vs file1, file1 vs file2
    compareDocxSemantic(original, file1, report.roundTrip1);
    compareDocxSemantic(file1, file2, report.roundTrip2);

    // Raw 解码数据对比
    report.rawDiff1 = compareRawDocuments(rawOriginal, rawFile1);
    report.rawDiff2 = compareRawDocuments(rawFile1, rawFile2);

  } catch (error: any) {
    report.roundTrip1.push({ category: 'FATAL', item: '', orig: '', restored: error.message, match: false });
  }

  return report;
}

// ============================================================
// XLSX 测试
// ============================================================

async function testXlsxFile(fileName: string): Promise<Report> {
  const report: Report = {
    file: fileName,
    fileSize: '',
    outputSize1: '',
    outputSize2: '',
    timings: { parse1: '', serialize1: '', parse2: '', serialize2: '', parse3: '' },
    roundTrip1: [],
    roundTrip2: [],
    rawDiff1: [],
    rawDiff2: [],
  };

  try {
    const buffer = await loadInputFile(fileName);
    report.fileSize = sizeKb(buffer.byteLength);

    const t0 = performance.now();
    const { raw: rawOriginal, semantic: original } = await parseXlsx(buffer);
    const t1 = performance.now();
    report.timings.parse1 = fmt(t1 - t0);
    report.original = original;
    report.rawOriginal = rawOriginal;

    const t2 = performance.now();
    const output1 = await serializeXlsx(original);
    const t3 = performance.now();
    report.timings.serialize1 = fmt(t3 - t2);
    report.outputSize1 = sizeKb(output1.byteLength);
    await saveOutputFile(fileName, output1);

    const t4 = performance.now();
    const { raw: rawFile1, semantic: file1 } = await parseXlsx(await loadOutputFile(fileName));
    const t5 = performance.now();
    report.timings.parse2 = fmt(t5 - t4);
    report.file1 = file1;
    report.rawFile1 = rawFile1;

    const t6 = performance.now();
    const output2 = await serializeXlsx(file1);
    const t7 = performance.now();
    report.timings.serialize2 = fmt(t7 - t6);
    report.outputSize2 = sizeKb(output2.byteLength);
    await saveOutputFile(fileName, output2);

    const t8 = performance.now();
    const { raw: rawFile2, semantic: file2 } = await parseXlsx(await loadOutputFile(fileName));
    const t9 = performance.now();
    report.timings.parse3 = fmt(t9 - t8);
    report.file2 = file2;
    report.rawFile2 = rawFile2;

    compareXlsxSemantic(original, file1, report.roundTrip1);
    compareXlsxSemantic(file1, file2, report.roundTrip2);

    // Raw 解码数据对比
    report.rawDiff1 = compareRawDocuments(rawOriginal, rawFile1);
    report.rawDiff2 = compareRawDocuments(rawFile1, rawFile2);

  } catch (error: any) {
    report.roundTrip1.push({ category: 'FATAL', item: '', orig: '', restored: error.message, match: false });
  }

  return report;
}

// ============================================================
// PPTX 测试
// ============================================================

async function testPptxFile(fileName: string): Promise<Report> {
  const report: Report = {
    file: fileName,
    fileSize: '',
    outputSize1: '',
    outputSize2: '',
    timings: { parse1: '', serialize1: '', parse2: '', serialize2: '', parse3: '' },
    roundTrip1: [],
    roundTrip2: [],
    rawDiff1: [],
    rawDiff2: [],
  };

  try {
    const buffer = await loadInputFile(fileName);
    report.fileSize = sizeKb(buffer.byteLength);

    const t0 = performance.now();
    const { raw: rawOriginal, semantic: original } = await parsePptx(buffer);
    const t1 = performance.now();
    report.timings.parse1 = fmt(t1 - t0);
    report.original = original;
    report.rawOriginal = rawOriginal;

    const t2 = performance.now();
    const output1 = await serializePptx(original);
    const t3 = performance.now();
    report.timings.serialize1 = fmt(t3 - t2);
    report.outputSize1 = sizeKb(output1.byteLength);
    await saveOutputFile(fileName, output1);

    const t4 = performance.now();
    const { raw: rawFile1, semantic: file1 } = await parsePptx(await loadOutputFile(fileName));
    const t5 = performance.now();
    report.timings.parse2 = fmt(t5 - t4);
    report.file1 = file1;
    report.rawFile1 = rawFile1;

    const t6 = performance.now();
    const output2 = await serializePptx(file1);
    const t7 = performance.now();
    report.timings.serialize2 = fmt(t7 - t6);
    report.outputSize2 = sizeKb(output2.byteLength);
    await saveOutputFile(fileName, output2);

    const t8 = performance.now();
    const { raw: rawFile2, semantic: file2 } = await parsePptx(await loadOutputFile(fileName));
    const t9 = performance.now();
    report.timings.parse3 = fmt(t9 - t8);
    report.file2 = file2;
    report.rawFile2 = rawFile2;

    comparePptxSemantic(original, file1, report.roundTrip1);
    comparePptxSemantic(file1, file2, report.roundTrip2);

    report.rawDiff1 = compareRawDocuments(rawOriginal, rawFile1);
    report.rawDiff2 = compareRawDocuments(rawFile1, rawFile2);

  } catch (error: any) {
    report.roundTrip1.push({ category: 'FATAL', item: '', orig: '', restored: error.message, match: false });
  }

  return report;
}

// ============================================================
// 还原度报告
// ============================================================

function printRawDiffReport(label: string, diffs: RawDiffResult[]) {
  const same = diffs.filter(d => d.status === 'same').length;
  const changed = diffs.filter(d => d.status === 'changed').length;
  const added = diffs.filter(d => d.status === 'added').length;
  const removed = diffs.filter(d => d.status === 'removed').length;
  const total = diffs.length;
  const restoreRate = total > 0 ? (same / total * 100) : 100;

  console.log(`\n  ┌─ ${label} ─────────────────────────────────`);
  console.log(`  │ XML Parts 还原度: ${same}/${total} (${restoreRate.toFixed(1)}%)${restoreRate === 100 ? ' ✓' : ''}`);

  if (changed + added + removed > 0) {
    console.log(`  │ 变更: ${changed} 修改, ${added} 新增, ${removed} 缺失`);
  }

  console.log(`  │`);
  console.log(`  │ Part                              状态`);
  console.log(`  │ ${'─'.repeat(55)}`);

  for (const d of diffs) {
    const icon = d.status === 'same' ? '✓' : d.status === 'changed' ? '~' : d.status === 'added' ? '+' : '-';
    const detail = d.details ? ` (${d.details})` : '';
    console.log(`  │ ${icon} ${d.part.padEnd(34)} ${d.status}${detail}`);
  }

  console.log(`  └──────────────────────────────────────────────`);
}

function printRestoreReport(label: string, details: CheckResult[]) {
  // 按类别分组
  const categories = new Map<string, { pass: number; fail: number; items: CheckResult[] }>();
  for (const d of details) {
    // category 格式: "Meta", "Block[0].para", "Block[0].run[1]", "Styles", "Comments" 等
    // 归类到顶层: Meta, Block, Styles, Comments, Revisions, Headers, Numbering, Footnotes, Section, Sheet, Workbook, FATAL
    let cat = d.category;
    if (cat.startsWith('Block[')) cat = 'Block';
    if (cat.startsWith('Comment[')) cat = 'Comments';
    if (cat.startsWith('Rev[')) cat = 'Revisions';
    if (cat.startsWith('Sheet[') || cat.includes('.cell[')) cat = 'Sheet';

    if (!categories.has(cat)) categories.set(cat, { pass: 0, fail: 0, items: [] });
    const group = categories.get(cat)!;
    if (d.match) group.pass++;
    else group.fail++;
    group.items.push(d);
  }

  const total = details.length;
  const passed = details.filter(d => d.match).length;
  const failed = total - passed;
  const rate = total > 0 ? (passed / total * 100) : 100;

  console.log(`\n  ┌─ ${label} ─────────────────────────────────`);
  console.log(`  │ 总还原度: ${passed}/${total} (${rate.toFixed(1)}%)${rate === 100 ? ' ✓' : ''}`);

  // 按类别展示
  console.log(`  │`);
  console.log(`  │ 类别             通过    失败    还原度`);
  console.log(`  │ ${'─'.repeat(50)}`);

  for (const [cat, group] of categories) {
    const catTotal = group.pass + group.fail;
    const catRate = catTotal > 0 ? (group.pass / catTotal * 100) : 100;
    const status = group.fail === 0 ? '✓' : '✗';
    const rateStr = catRate === 100 ? '100%' : `${catRate.toFixed(1)}%`;
    console.log(`  │ ${status} ${cat.padEnd(16)} ${String(group.pass).padStart(5)}  ${String(group.fail).padStart(5)}    ${rateStr.padStart(6)}`);
  }

  // 展示不匹配详情
  const issues = details.filter(d => !d.match);
  if (issues.length > 0) {
    console.log(`  │`);
    console.log(`  │ ── 不匹配详情 ──`);

    // 按类别分组展示
    const issueGroups = new Map<string, CheckResult[]>();
    for (const issue of issues) {
      let cat = issue.category;
      if (cat.startsWith('Block[')) {
        const blockIdx = cat.match(/Block\[(\d+)\]/)?.[1] || '?';
        cat = `Block[${blockIdx}]`;
      }
      if (!issueGroups.has(cat)) issueGroups.set(cat, []);
      issueGroups.get(cat)!.push(issue);
    }

    for (const [cat, catIssues] of issueGroups) {
      console.log(`  │`);
      console.log(`  │ ${cat}:`);
      for (const issue of catIssues.slice(0, 15)) {
        const origShort = issue.orig.length > 40 ? issue.orig.slice(0, 37) + '...' : issue.orig;
        const restShort = issue.restored.length > 40 ? issue.restored.slice(0, 37) + '...' : issue.restored;
        console.log(`  │   ${issue.item}: ${origShort} → ${restShort}`);
      }
      if (catIssues.length > 15) {
        console.log(`  │   ... 还有 ${catIssues.length - 15} 处`);
      }
    }
  }

  console.log(`  └──────────────────────────────────────────────`);
}

function getReportType(report: Report): 'docx' | 'pptx' | 'xlsx' {
  if (report.original && 'body' in report.original) return 'docx';
  if (report.original && 'slides' in report.original) return 'pptx';
  return 'xlsx';
}

function printSemantic(label: string, data: DocxDocument | PptxPresentation | XlsxWorkbook, type: 'docx' | 'pptx' | 'xlsx') {
  if (type === 'docx') printDocxSemantic(label, data as DocxDocument);
  else if (type === 'pptx') printPptxSemantic(label, data as PptxPresentation);
  else printXlsxSemantic(label, data as XlsxWorkbook);
}

function printReport(report: Report) {
  const type = getReportType(report);

  if (report.original) printSemantic('第 1 次解析 (原始文件)', report.original, type);
  if (report.file1) printSemantic('第 2 次解析 (输出文件 1)', report.file1, type);
  if (report.file2) printSemantic('第 3 次解析 (输出文件 2)', report.file2, type);

  // 耗时
  console.log(`\n  ┌─ 耗时 ────────────────────────────────────`);
  console.log(`  │ ${report.file} (${report.fileSize})`);
  console.log(`  │ 解析1: ${report.timings.parse1}  序列化1: ${report.timings.serialize1}  → ${report.outputSize1}`);
  console.log(`  │ 解析2: ${report.timings.parse2}  序列化2: ${report.timings.serialize2}  → ${report.outputSize2}`);
  console.log(`  │ 解析3: ${report.timings.parse3}`);
  console.log(`  └──────────────────────────────────────────`);

  // Raw 解码数据对比
  printRawDiffReport('源文件 vs 输出文件 1: XML Parts 解码对比', report.rawDiff1);
  printRawDiffReport('输出文件 1 vs 输出文件 2: XML Parts 解码对比', report.rawDiff2);

  // Round-trip 1: 原始文件解码 vs 输出文件解码
  printRestoreReport('源文件 vs 输出文件 1: 语义解码还原度', report.roundTrip1);

  // Round-trip 2: 输出文件 1 解码 vs 输出文件 2 解码
  printRestoreReport('输出文件 1 vs 输出文件 2: 语义解码还原度', report.roundTrip2);
}

// ============================================================
// 主入口
// ============================================================

const DOCX_FILES = [
  '(2025)鄂0192行初0011号_4.docx',
  '判决书示例.docx',
  'demo_list.docx',
  'demo_long.docx',
  'demo_test.docx',
  '批注demo.docx',
  '修订demo.docx',
  'document_1769426663776.docx',
];

const PPTX_FILES = [
  'test.pptx',
  'demo-01.pptx',
];

const XLSX_FILES = [
  'demo_xlsx.xlsx',
  'demo-02.xlsx',
];

async function main() {
  console.log('╔══════════════════════════════════════════════════════════════════════╗');
  console.log('║      Office Meta Parser · 文件级双重 Round-trip 测试报告          ║');
  console.log('╚══════════════════════════════════════════════════════════════════════╝');
  console.log(`  时间: ${new Date().toLocaleString('zh-CN')}`);
  console.log(`  文件数: DOCX ${DOCX_FILES.length} + PPTX ${PPTX_FILES.length} + XLSX ${XLSX_FILES.length}`);
  console.log(`  输出目录: ${OUTPUT_DIR}`);
  console.log(`  流程: 原始文件 → 解析 → 序列化 → 写文件 → 读文件 → 解析 → 序列化 → 写文件 → 读文件 → 解析`);

  const reports: Report[] = [];

  // DOCX 测试
  console.log('\n▶ DOCX 文件测试');
  for (const file of DOCX_FILES) {
    process.stdout.write(`  ${file} ...`);
    const report = await testDocxFile(file);
    reports.push(report);
    const r1issues = report.roundTrip1.filter(d => !d.match).length;
    const r2issues = report.roundTrip2.filter(d => !d.match).length;
    const raw2issues = report.rawDiff2.filter(d => d.status !== 'same').length;
    const status = r1issues === 0 && r2issues === 0 && raw2issues === 0 ? 'PASS' : 'FAIL';
    console.log(` ${status} (语义:${r1issues}/${r2issues} XML确定性:${raw2issues})`);
  }

  // PPTX 测试
  console.log('\n▶ PPTX 文件测试');
  for (const file of PPTX_FILES) {
    process.stdout.write(`  ${file} ...`);
    const report = await testPptxFile(file);
    reports.push(report);
    const r1issues = report.roundTrip1.filter(d => !d.match).length;
    const r2issues = report.roundTrip2.filter(d => !d.match).length;
    const raw2issues = report.rawDiff2.filter(d => d.status !== 'same').length;
    const status = r1issues === 0 && r2issues === 0 && raw2issues === 0 ? 'PASS' : 'FAIL';
    console.log(` ${status} (语义:${r1issues}/${r2issues} XML确定性:${raw2issues})`);
  }

  // XLSX 测试
  console.log('\n▶ XLSX 文件测试');
  for (const file of XLSX_FILES) {
    process.stdout.write(`  ${file} ...`);
    const report = await testXlsxFile(file);
    reports.push(report);
    const r1issues = report.roundTrip1.filter(d => !d.match).length;
    const r2issues = report.roundTrip2.filter(d => !d.match).length;
    const raw2issues = report.rawDiff2.filter(d => d.status !== 'same').length;
    const status = r1issues === 0 && r2issues === 0 && raw2issues === 0 ? 'PASS' : 'FAIL';
    console.log(` ${status} (语义:${r1issues}/${r2issues} XML确定性:${raw2issues})`);
  }

  // 详细报告
  for (const report of reports) {
    printReport(report);
  }

  // 总结
  console.log(`\n${'═'.repeat(70)}`);
  console.log('  总结');
  console.log(`${'═'.repeat(70)}`);

  const allDetails1 = reports.flatMap(r => r.roundTrip1);
  const allDetails2 = reports.flatMap(r => r.roundTrip2);
  const passed1 = allDetails1.filter(d => d.match).length;
  const failed1 = allDetails1.filter(d => !d.match).length;
  const passed2 = allDetails2.filter(d => d.match).length;
  const failed2 = allDetails2.filter(d => !d.match).length;
  const rate1 = allDetails1.length > 0 ? (passed1 / allDetails1.length * 100) : 100;
  const rate2 = allDetails2.length > 0 ? (passed2 / allDetails2.length * 100) : 100;

  console.log(`  Round-trip 1 (源文件 vs 输出文件1): ${passed1}/${allDetails1.length} (${rate1.toFixed(1)}%)`);
  console.log(`  Round-trip 2 (输出文件1 vs 输出文件2): ${passed2}/${allDetails2.length} (${rate2.toFixed(1)}%)`);

  // Raw 对比统计
  const allRaw1 = reports.flatMap(r => r.rawDiff1);
  const allRaw2 = reports.flatMap(r => r.rawDiff2);
  const rawSame1 = allRaw1.filter(d => d.status === 'same').length;
  const rawSame2 = allRaw2.filter(d => d.status === 'same').length;
  const rawRate1 = allRaw1.length > 0 ? (rawSame1 / allRaw1.length * 100) : 100;
  const rawRate2 = allRaw2.length > 0 ? (rawSame2 / allRaw2.length * 100) : 100;

  console.log(`\n  XML Parts 还原度:`);
  console.log(`  Round-trip 1: ${rawSame1}/${allRaw1.length} (${rawRate1.toFixed(1)}%)`);
  console.log(`  Round-trip 2: ${rawSame2}/${allRaw2.length} (${rawRate2.toFixed(1)}%)`);

  // 每个文件的还原度
  console.log(`\n  文件还原度:`);
  for (const r of reports) {
    const r1p = r.roundTrip1.filter(d => d.match).length;
    const r1t = r.roundTrip1.length;
    const r2p = r.roundTrip2.filter(d => d.match).length;
    const r2t = r.roundTrip2.length;
    const r1rate = r1t > 0 ? (r1p / r1t * 100) : 100;
    const r2rate = r2t > 0 ? (r2p / r2t * 100) : 100;
    const raw1same = r.rawDiff1.filter(d => d.status === 'same').length;
    const raw1total = r.rawDiff1.length;
    const raw2same = r.rawDiff2.filter(d => d.status === 'same').length;
    const raw2total = r.rawDiff2.length;
    const raw1rate = raw1total > 0 ? (raw1same / raw1total * 100) : 100;
    const raw2rate = raw2total > 0 ? (raw2same / raw2total * 100) : 100;
    // 语义 100% + XML 确定性 100% = PASS
    const semanticOk = r1rate === 100 && r2rate === 100;
    const xmlDeterministic = raw2rate === 100;
    const status = semanticOk && xmlDeterministic ? 'PASS' : 'FAIL';
    console.log(`    ${status} ${r.file.padEnd(40)} 语义: ${r1rate.toFixed(1)}%/${r2rate.toFixed(1)}%  XML确定性: ${raw2rate.toFixed(1)}%`);
  }

  const filePassed = reports.filter(r => {
    const r1ok = r.roundTrip1.every(d => d.match);
    const r2ok = r.roundTrip2.every(d => d.match);
    const xmlDet = r.rawDiff2.every(d => d.status === 'same');
    return r1ok && r2ok && xmlDet;
  }).length;
  const fileFailed = reports.length - filePassed;
  console.log(`\n  文件: ${filePassed} 通过 / ${fileFailed} 失败 / ${reports.length} 总计`);

  // 问题汇总
  const issueProps = new Map<string, number>();
  for (const d of [...allDetails1, ...allDetails2]) {
    if (!d.match) {
      issueProps.set(d.item, (issueProps.get(d.item) || 0) + 1);
    }
  }
  if (issueProps.size > 0) {
    console.log('\n  问题属性分布:');
    const sorted = [...issueProps.entries()].sort((a, b) => b[1] - a[1]);
    for (const [prop, count] of sorted) {
      console.log(`    ${prop.padEnd(30)} ${count} 处`);
    }
  }

  console.log(`\n${'═'.repeat(70)}`);

  if (fileFailed > 0) process.exit(1);
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
