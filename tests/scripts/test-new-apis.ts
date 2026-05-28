#!/usr/bin/env tsx
/**
 * 新 API 测试脚本
 * 测试 create* / update*Meta / detectFormat / validate / toBuffer / ValidationError
 */

import { mkdirSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { createDocx, loadDocx, saveDocx, validateDocx, serializeDocx, docx, updateDocxTitle, updateDocxCreator, updateDocxSubject, updateDocxDescription, updateDocxKeywords, updateDocxCategory, updateDocxLastModifiedBy, toDocxJSON, toDocxJSONString, saveDocxJSON, addComment, removeComment, listComments, getCommentText, markCommentDone, markCommentUndone, markInsert, markDelete, addFormatChange, clearRevision, listRevisions, hasPendingRevisions, acceptAllInserts, acceptAllDeletes, rejectAllInserts, rejectAllDeletes } from '../../src/docx/index.js';
import { createXlsx, loadXlsx, saveXlsx, validateXlsx, serializeXlsx, xlsx, updateXlsxTitle, updateXlsxCreator, toXlsxJSON, toXlsxJSONString, saveXlsxJSON, addComment as addXlsxComment, removeComment as removeXlsxComment, listComments as listXlsxComments, getCommentText as getXlsxCommentText, updateComment as updateXlsxComment } from '../../src/xlsx/index.js';
import { createPptx, loadPptx, savePptx, validatePptx, serializePptx, pptx, updatePptxTitle, updatePptxCreator, toPptxJSON, toPptxJSONString, savePptxJSON, addComment as addPptxComment, removeComment as removePptxComment, listComments as listPptxComments, listSlideComments, getCommentText as getPptxCommentText } from '../../src/pptx/index.js';
import { detectFormat, validate, toBuffer, toJSON, toJSONString, saveToJSON, throwOnError, ValidationError, FormatError } from '../../src/core/index.js';
import type { ValidationIssue } from '../../src/core/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const OUTPUT_DIR = join(__dirname, '..', 'output', 'new-apis');

mkdirSync(OUTPUT_DIR, { recursive: true });

let passed = 0;
let failed = 0;

function assert(condition: boolean, message: string) {
  if (condition) {
    passed++;
    console.log(`  ✓ ${message}`);
  } else {
    failed++;
    console.error(`  ✗ ${message}`);
  }
}

function assertEqual(actual: any, expected: any, message: string) {
  const ok = actual === expected;
  if (!ok) {
    console.error(`    expected: ${expected}, actual: ${actual}`);
  }
  assert(ok, message);
}

// ============================================================
// 1. create* 空文档工厂
// ============================================================

async function testCreateDocx() {
  console.log('\n--- createDocx ---');
  const doc = createDocx({ title: '测试文档', creator: '张三' });

  assertEqual(doc.meta.title, '测试文档', 'meta.title');
  assertEqual(doc.meta.creator, '张三', 'meta.creator');
  assert(doc.meta.created !== undefined, 'meta.created 自动填充');
  assertEqual(doc.body.blocks.length, 0, 'body.blocks 为空数组');
  assertEqual(doc.styles.paragraphStyles.length, 0, 'styles 初始化');

  // 添加内容后可正常序列化
  doc.body.blocks.push({
    type: 'paragraph',
    runs: [{ text: 'Hello World', bold: true, fontSize: 24, color: 'FF0000' }],
  });

  const output = await serializeDocx(doc);
  assert(output.byteLength > 0, 'serializeDocx 产出非空');

  const issues = validateDocx(doc);
  assert(issues.filter(i => i.level === 'error').length === 0, 'validateDocx 无 error');

  // round-trip
  await saveDocx(doc, join(OUTPUT_DIR, 'create-docx.docx'));
  const loaded = await loadDocx(join(OUTPUT_DIR, 'create-docx.docx'));
  assertEqual(loaded.semantic.meta.title, '测试文档', 'round-trip meta.title');
  assertEqual(loaded.semantic.body.blocks.length, 1, 'round-trip body.blocks');
}

async function testCreateXlsx() {
  console.log('\n--- createXlsx ---');
  const wb = createXlsx({ title: '测试表格', creator: '李四', sheetName: '数据' });

  assertEqual(wb.meta.title, '测试表格', 'meta.title');
  assertEqual(wb.sheets.length, 1, '默认 1 个 sheet');
  assertEqual(wb.sheets[0].name, '数据', '自定义 sheetName');
  assertEqual(wb.sharedStrings.length, 0, 'sharedStrings 初始化');
  assert(wb.styles.fonts.length >= 0, 'styles 初始化');

  // 添加数据
  wb.sheets[0].cells.push([
    { value: '姓名', type: 'string' },
    { value: '年龄', type: 'string' },
  ]);
  wb.sheets[0].cells.push([
    { value: '张三', type: 'string' },
    { value: 28, type: 'number' },
  ]);

  const output = await serializeXlsx(wb);
  assert(output.byteLength > 0, 'serializeXlsx 产出非空');

  await saveXlsx(wb, join(OUTPUT_DIR, 'create-xlsx.xlsx'));
  const loaded = await loadXlsx(join(OUTPUT_DIR, 'create-xlsx.xlsx'));
  assertEqual(loaded.semantic.meta.title, '测试表格', 'round-trip meta.title');
  assertEqual(loaded.semantic.sheets[0].name, '数据', 'round-trip sheet name');
}

async function testCreatePptx() {
  console.log('\n--- createPptx ---');
  const pres = createPptx({ title: '测试PPT', creator: '王五' });

  assertEqual(pres.meta.title, '测试PPT', 'meta.title');
  assert(pres.slides.length >= 1, '至少 1 张幻灯片');
  assert(pres.masters.length >= 1, '至少 1 个 master');
  assert(pres.layouts.length >= 1, '至少 1 个 layout');
  assert(pres.theme !== undefined, '自带 theme');
  assert(pres.slideSize !== undefined, '自带 slideSize');

  // 添加内容
  pres.slides[0].elements.push({
    type: 'text',
    content: '标题',
    position: { x: 0, y: 0, width: 9144000, height: 1000000 },
    paragraphs: [{ runs: [{ text: '标题文本' }] }],
  });

  const output = await serializePptx(pres);
  assert(output.byteLength > 0, 'serializePptx 产出非空');

  await savePptx(pres, join(OUTPUT_DIR, 'create-pptx.pptx'));
  const loaded = await loadPptx(join(OUTPUT_DIR, 'create-pptx.pptx'));
  assertEqual(loaded.semantic.meta.title, '测试PPT', 'round-trip meta.title');
  assertEqual(loaded.semantic.slides.length, 1, 'round-trip slides');
}

// ============================================================
// 2. update*Meta 细粒度 API
// ============================================================

async function testUpdateDocxMeta() {
  console.log('\n--- updateDocxMeta ---');
  const doc = createDocx({ title: '原始' });

  // 独立函数
  updateDocxTitle(doc, '新标题');
  assertEqual(doc.meta.title, '新标题', 'updateDocxTitle');

  updateDocxCreator(doc, '张三');
  assertEqual(doc.meta.creator, '张三', 'updateDocxCreator');

  updateDocxSubject(doc, '主题');
  assertEqual(doc.meta.subject, '主题', 'updateDocxSubject');

  updateDocxDescription(doc, '描述');
  assertEqual(doc.meta.description, '描述', 'updateDocxDescription');

  updateDocxKeywords(doc, '关键词');
  assertEqual(doc.meta.keywords, '关键词', 'updateDocxKeywords');

  updateDocxCategory(doc, '分类');
  assertEqual(doc.meta.category, '分类', 'updateDocxCategory');

  updateDocxLastModifiedBy(doc, '修改者');
  assertEqual(doc.meta.lastModifiedBy, '修改者', 'updateDocxLastModifiedBy');

  // 命名空间
  docx.updateTitle(doc, '命名空间标题');
  assertEqual(doc.meta.title, '命名空间标题', 'docx.updateTitle');

  docx.updateCreator(doc, '命名空间创建者');
  assertEqual(doc.meta.creator, '命名空间创建者', 'docx.updateCreator');

  // modified 自动更新
  const before = doc.meta.modified;
  await new Promise(r => setTimeout(r, 10));
  updateDocxTitle(doc, '触发modified');
  assert(doc.meta.modified !== before, 'modified 自动更新');
}

async function testUpdateXlsxMeta() {
  console.log('\n--- updateXlsxMeta ---');
  const wb = createXlsx({ title: '原始' });

  updateXlsxTitle(wb, '新表格');
  assertEqual(wb.meta.title, '新表格', 'updateXlsxTitle');

  xlsx.updateCreator(wb, '李四');
  assertEqual(wb.meta.creator, '李四', 'xlsx.updateCreator');
}

async function testUpdatePptxMeta() {
  console.log('\n--- updatePptxMeta ---');
  const pres = createPptx({ title: '原始' });

  updatePptxTitle(pres, '新PPT');
  assertEqual(pres.meta.title, '新PPT', 'updatePptxTitle');

  pptx.updateCreator(pres, '王五');
  assertEqual(pres.meta.creator, '王五', 'pptx.updateCreator');
}

// ============================================================
// 3. detectFormat 格式检测
// ============================================================

async function testDetectFormat() {
  console.log('\n--- detectFormat ---');

  const doc = createDocx({ title: 'test' });
  doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'x' }] });
  const docAb = await serializeDocx(doc);
  assertEqual(await detectFormat(docAb), 'docx', 'detect docx');

  const wb = createXlsx({ title: 'test' });
  const xlsxAb = await serializeXlsx(wb);
  assertEqual(await detectFormat(xlsxAb), 'xlsx', 'detect xlsx');

  const pres = createPptx({ title: 'test' });
  const pptxAb = await serializePptx(pres);
  assertEqual(await detectFormat(pptxAb), 'pptx', 'detect pptx');

  // 非 ZIP 文件
  const garbage = new TextEncoder().encode('not a zip').buffer;
  assertEqual(await detectFormat(garbage), null, 'detect null for garbage');
}

// ============================================================
// 4. validate(buffer) 统一校验
// ============================================================

async function testValidateBuffer() {
  console.log('\n--- validate(buffer) ---');

  const doc = createDocx({ title: 'test' });
  doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'x' }] });
  const ab = await serializeDocx(doc);

  const result = await validate(ab);
  assertEqual(result.format, 'docx', 'result.format');
  assertEqual(result.valid, true, 'result.valid');
  assert(Array.isArray(result.issues), 'result.issues 是数组');

  // 非 Office 文件
  const garbage = new TextEncoder().encode('not a zip').buffer;
  const bad = await validate(garbage);
  assertEqual(bad.format, null, 'garbage format is null');
  assertEqual(bad.valid, false, 'garbage valid is false');
  assert(bad.issues.length > 0, 'garbage has issues');
}

// ============================================================
// 5. toBuffer 转换
// ============================================================

async function testToBuffer() {
  console.log('\n--- toBuffer ---');

  const doc = createDocx({ title: 'test' });
  doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'x' }] });
  const ab = await serializeDocx(doc);

  const buf = toBuffer(ab);
  assert(Buffer.isBuffer(buf), '返回 Node.js Buffer');
  assertEqual(buf.length, ab.byteLength, '长度一致');
}

// ============================================================
// 6. ValidationError / FormatError
// ============================================================

async function testErrorTypes() {
  console.log('\n--- ValidationError / FormatError ---');

  // ValidationError
  try {
    throwOnError([
      { level: 'error', path: 'body', message: 'blocks is empty' },
      { level: 'warning', path: 'meta', message: 'no title' },
    ]);
    assert(false, '应抛出 ValidationError');
  } catch (e: any) {
    assert(e instanceof ValidationError, 'catch instanceof ValidationError');
    assertEqual(e.name, 'ValidationError', 'error.name');
    assertEqual(e.issues.length, 1, 'issues 只包含 error 级别');
    assertEqual(e.issues[0].path, 'body', 'issues[0].path');
  }

  // FormatError
  const err = new FormatError('不支持的格式', 'pdf');
  assertEqual(err.name, 'FormatError', 'FormatError.name');
  assertEqual(err.format, 'pdf', 'FormatError.format');
  assert(err.message.includes('不支持'), 'FormatError.message');

  // warning 不抛异常
  try {
    throwOnError([{ level: 'warning', path: 'x', message: 'warn' }]);
    assert(true, 'warning 不抛异常');
  } catch {
    assert(false, 'warning 不应抛异常');
  }
}

// ============================================================
// 7. 完整业务流程模拟
// ============================================================

async function testFullWorkflow() {
  console.log('\n--- 完整业务流程 ---');

  // 1. 从零创建
  const doc = createDocx({ title: '月度报告', creator: '系统' });

  // 2. 设置元数据
  docx.updateTitle(doc, '2024年5月月度报告');
  docx.updateCategory(doc, '工作报告');

  // 3. 添加内容
  doc.body.blocks.push(
    { type: 'paragraph', runs: [{ text: '一、概述', bold: true, fontSize: 28 }] },
    { type: 'paragraph', runs: [{ text: '本月完成了 office-meta-parser 的 API 设计与实现。' }] },
    { type: 'paragraph', runs: [{ text: '二、成果', bold: true, fontSize: 28 }] },
    { type: 'paragraph', runs: [
      { text: '新增' },
      { text: '6 项', bold: true, color: 'FF0000' },
      { text: '核心 API。' },
    ]},
  );

  // 4. 校验
  const issues = validateDocx(doc);
  assertEqual(issues.filter(i => i.level === 'error').length, 0, '校验通过');

  // 5. 序列化 + 格式检测
  const ab = await serializeDocx(doc);
  const format = await detectFormat(ab);
  assertEqual(format, 'docx', 'detectFormat 确认格式');

  // 6. toBuffer
  const buf = toBuffer(ab);
  assert(Buffer.isBuffer(buf), 'toBuffer 转换');

  // 7. 保存 + 加载 + 验证
  const path = join(OUTPUT_DIR, 'workflow-report.docx');
  await saveDocx(doc, path);
  const loaded = await loadDocx(path);
  assertEqual(loaded.semantic.meta.title, '2024年5月月度报告', '加载后标题正确');
  assertEqual(loaded.semantic.body.blocks.length, 4, '加载后内容正确');

  // 8. 修改元数据后重新保存
  docx.updateTitle(doc, '2024年5月月度报告（修订版）');
  await saveDocx(doc, join(OUTPUT_DIR, 'workflow-report-v2.docx'));
  const loaded2 = await loadDocx(join(OUTPUT_DIR, 'workflow-report-v2.docx'));
  assertEqual(loaded2.semantic.meta.title, '2024年5月月度报告（修订版）', '修订后标题正确');
}

// ============================================================
// 8. toJSON / toJSONString / saveToJSON
// ============================================================

async function testToJSON() {
  console.log('\n--- toJSON ---');

  // DOCX
  const doc = createDocx({ title: 'JSON测试', creator: '张三' });
  doc.body.blocks.push({ type: 'paragraph', runs: [{ text: '正文' }] });

  const docJson = toDocxJSON(doc);
  assertEqual(docJson.meta.title, 'JSON测试', 'toDocxJSON meta.title');
  assertEqual(docJson.body.blocks.length, 1, 'toDocxJSON body.blocks');
  assert(typeof docJson === 'object', 'toDocxJSON 返回对象');

  const docStr = toDocxJSONString(doc, 2);
  assert(docStr.includes('"title": "JSON测试"'), 'toDocxJSONString 包含字段');
  const parsed = JSON.parse(docStr);
  assertEqual(parsed.meta.creator, '张三', 'JSON.parse 后 creator 正确');

  await saveDocxJSON(doc, join(OUTPUT_DIR, 'docx-meta.json'));
  assert(existsSync(join(OUTPUT_DIR, 'docx-meta.json')), 'saveDocxJSON 写入文件');

  // 命名空间方式
  const docJson2 = docx.toJSON(doc);
  assertEqual(docJson2.meta.title, 'JSON测试', 'docx.toJSON');
  const docStr2 = docx.toJSONString(doc);
  assert(docStr2.length > 0, 'docx.toJSONString');

  // XLSX
  const wb = createXlsx({ title: '表格JSON', sheetName: '数据' });
  wb.sheets[0].cells.push([{ value: 'A1', type: 'string' }]);

  const wbJson = toXlsxJSON(wb);
  assertEqual(wbJson.meta.title, '表格JSON', 'toXlsxJSON meta.title');
  assertEqual(wbJson.sheets.length, 1, 'toXlsxJSON sheets');

  const wbStr = toXlsxJSONString(wb);
  assert(wbStr.includes('表格JSON'), 'toXlsxJSONString 包含字段');

  await saveXlsxJSON(wb, join(OUTPUT_DIR, 'xlsx-meta.json'));
  assert(existsSync(join(OUTPUT_DIR, 'xlsx-meta.json')), 'saveXlsxJSON 写入文件');

  const wbJson2 = xlsx.toJSON(wb);
  assertEqual(wbJson2.meta.title, '表格JSON', 'xlsx.toJSON');

  // PPTX
  const pres = createPptx({ title: 'PPTJSON' });
  const presJson = toPptxJSON(pres);
  assertEqual(presJson.meta.title, 'PPTJSON', 'toPptxJSON meta.title');
  assert(presJson.masters.length >= 1, 'toPptxJSON masters');

  const presStr = toPptxJSONString(pres);
  assert(presStr.includes('PPTJSON'), 'toPptxJSONString 包含字段');

  await savePptxJSON(pres, join(OUTPUT_DIR, 'pptx-meta.json'));
  assert(existsSync(join(OUTPUT_DIR, 'pptx-meta.json')), 'savePptxJSON 写入文件');

  const presJson2 = pptx.toJSON(pres);
  assertEqual(presJson2.meta.title, 'PPTJSON', 'pptx.toJSON');

  // core 泛型
  const coreJson = toJSON(doc);
  assertEqual(coreJson.meta.title, 'JSON测试', 'core toJSON');
  const coreStr = toJSONString(doc, 0);
  assert(coreStr.length > 0 && !coreStr.includes('\n'), 'core toJSONString space=0 无换行');

  await saveToJSON(doc, join(OUTPUT_DIR, 'core-meta.json'));
  assert(existsSync(join(OUTPUT_DIR, 'core-meta.json')), 'core saveToJSON');

  // rawXmlParts 被剥离
  const docWithRaw = { ...doc, rawXmlParts: new Map([['a', 'b']]) };
  const cleanJson = toJSON(docWithRaw);
  assert(!('rawXmlParts' in cleanJson), 'rawXmlParts 被剥离');
}

// ============================================================
// 9. 批注 (Comments)
// ============================================================

async function testComments() {
  console.log('\n--- Comments ---');

  const doc = createDocx({ title: '批注测试' });
  const run1 = { text: '第一段内容', bold: true };
  const run2 = { text: '第二段内容' };
  doc.body.blocks.push(
    { type: 'paragraph', runs: [run1] },
    { type: 'paragraph', runs: [run2] },
  );

  // addComment
  const c1 = addComment(doc, run1, '张三', '这里需要修改');
  assert(c1.id !== undefined, 'addComment 返回带 id');
  assertEqual(run1.commentId, c1.id, 'run.commentId 已设置');
  assertEqual(doc.comments!.length, 1, 'doc.comments 有 1 条');

  const c2 = addComment(doc, run2, '李四', '这里也需要改', '2024-01-01T00:00:00Z');
  assertEqual(doc.comments!.length, 2, 'doc.comments 有 2 条');

  // getCommentText
  const text = getCommentText(doc, c1.id);
  assertEqual(text, '这里需要修改', 'getCommentText 正确');

  // listComments
  const list = listComments(doc);
  assertEqual(list.length, 2, 'listComments 返回 2 条');
  assertEqual(list[0].comment.author, '张三', 'listComments[0].author');

  // markCommentDone
  assert(markCommentDone(doc, c1.id), 'markCommentDone 成功');
  assert(doc.commentExts!.some(e => e.paraId === c1.id && e.done), 'commentExts.done=true');

  // markCommentUndone
  assert(markCommentUndone(doc, c1.id), 'markCommentUndone 成功');
  assert(doc.commentExts!.some(e => e.paraId === c1.id && !e.done), 'commentExts.done=false');

  // removeComment
  assert(removeComment(doc, c2.id), 'removeComment 成功');
  assertEqual(doc.comments!.length, 1, '移除后 1 条');
  assertEqual(run2.commentId, undefined, 'run.commentId 已清除');

  // removeComment 不存在的 ID
  assert(!removeComment(doc, 'nonexistent'), 'removeComment 不存在返回 false');

  // 序列化 round-trip
  const ab = await serializeDocx(doc);
  assert(ab.byteLength > 0, 'serializeDocx 产出非空');
  const tmpPath = join(OUTPUT_DIR, 'test-comments.docx');
  await saveDocx(doc, tmpPath);
  const loaded = await loadDocx(tmpPath);
  assert(loaded.semantic.comments!.length >= 1, 'round-trip 保留批注');

  // validateDocx 通过
  const issues = validateDocx(doc);
  assertEqual(issues.filter(i => i.level === 'error').length, 0, '带批注的文档校验通过');
}

// ============================================================
// 10. 修订 (Revisions)
// ============================================================

async function testRevisions() {
  console.log('\n--- Revisions ---');

  const doc = createDocx({ title: '修订测试' });
  const run1 = { text: '原始文本' };
  const run2 = { text: '新增文本' };
  const run3 = { text: '要删除的文本' };
  doc.body.blocks.push(
    { type: 'paragraph', runs: [run1, run2, run3] },
  );

  // markInsert
  markInsert(run2, '张三', '2024-01-01T00:00:00Z');
  assertEqual(run2.revisionType, 'insert', 'markInsert 设置 revisionType');
  assertEqual(run2.revisionAuthor, '张三', 'markInsert 设置 revisionAuthor');

  // markDelete
  markDelete(run3, '张三');
  assertEqual(run3.revisionType, 'delete', 'markDelete 设置 revisionType');

  // hasPendingRevisions
  assert(hasPendingRevisions(doc), 'hasPendingRevisions = true');

  // listRevisions
  const revisions = listRevisions(doc);
  assertEqual(revisions.length, 2, 'listRevisions 返回 2 条');
  assertEqual(revisions[0].revision.type, 'insert', 'revisions[0].type');
  assertEqual(revisions[1].revision.type, 'delete', 'revisions[1].type');

  // addFormatChange
  addFormatChange(doc, '李四');
  assertEqual(doc.trackChanges!.length, 1, 'addFormatChange 添加记录');
  assertEqual(doc.trackChanges![0].type, 'formatChange', 'trackChanges.type');

  // clearRevision
  clearRevision(run2);
  assertEqual(run2.revisionType, undefined, 'clearRevision 清除 revisionType');
  assertEqual(run2.revisionAuthor, undefined, 'clearRevision 清除 revisionAuthor');

  // acceptAllInserts
  markInsert(run1, '王五');
  const accepted = acceptAllInserts(doc);
  assert(accepted >= 1, 'acceptAllInserts 处理了修订');
  assertEqual(run1.revisionType, undefined, 'acceptAllInserts 清除标记');

  // acceptAllDeletes
  markDelete(run3, '赵六');
  const docBefore = doc.body.blocks[0] as any;
  const runCountBefore = docBefore.runs.length;
  const deleted = acceptAllDeletes(doc);
  assert(deleted >= 1, 'acceptAllDeletes 处理了修订');
  assert((doc.body.blocks[0] as any).runs.length < runCountBefore, 'acceptAllDeletes 移除了 run');

  // 序列化 round-trip
  const doc2 = createDocx({ title: '修订round-trip' });
  const r = { text: '带修订的文本' };
  markInsert(r, '张三');
  doc2.body.blocks.push({ type: 'paragraph', runs: [r] });
  const tmpPath = join(OUTPUT_DIR, 'test-revisions.docx');
  await saveDocx(doc2, tmpPath);
  const loaded = await loadDocx(tmpPath);
  const loadedRun = (loaded.semantic.body.blocks[0] as any).runs[0];
  assertEqual(loadedRun.revisionType, 'insert', 'round-trip 保留 revisionType');
  assertEqual(loadedRun.revisionAuthor, '张三', 'round-trip 保留 revisionAuthor');

  // rejectAllDeletes
  const doc3 = createDocx({ title: 'reject测试' });
  const r1 = { text: '保留' };
  const r2 = { text: '删除' };
  markDelete(r2, '张三');
  doc3.body.blocks.push({ type: 'paragraph', runs: [r1, r2] });
  const rejected = rejectAllDeletes(doc3);
  assert(rejected >= 1, 'rejectAllDeletes 处理了修订');
  assertEqual(r2.revisionType, undefined, 'rejectAllDeletes 移除标记，保留文本');
  assertEqual((doc3.body.blocks[0] as any).runs.length, 2, 'rejectAllDeletes 不移除 run');
}

// ============================================================
// 11. XLSX 批注
// ============================================================

async function testXlsxComments() {
  console.log('\n--- XLSX Comments ---');

  const wb = createXlsx({ title: '批注测试' });

  // addComment
  const c1 = addXlsxComment(wb, 0, 'A1', '张三', '需要修改');
  assertEqual(c1.ref, 'A1', 'addComment ref 正确');
  assertEqual(c1.text, '需要修改', 'addComment text 正确');

  const c2 = addXlsxComment(wb, 0, 'B2', '李四', '数据有误');
  assertEqual(wb.sheets[0].comments!.length, 2, 'sheet 有 2 条批注');

  // getCommentText
  assertEqual(getXlsxCommentText(wb, 0, 'A1'), '需要修改', 'getCommentText 正确');

  // listComments
  const all = listXlsxComments(wb);
  assertEqual(all.length, 2, 'listComments 返回 2 条');

  // updateComment
  assert(updateXlsxComment(wb, 0, 'A1', '已修改'), 'updateComment 成功');
  assertEqual(getXlsxCommentText(wb, 0, 'A1'), '已修改', '更新后文本正确');

  // removeComment
  assert(removeXlsxComment(wb, 0, 'B2'), 'removeComment 成功');
  assertEqual(wb.sheets[0].comments!.length, 1, '移除后 1 条');

  // round-trip
  const tmpPath = join(OUTPUT_DIR, 'test-xlsx-comments.xlsx');
  await saveXlsx(wb, tmpPath);
  const loaded = await loadXlsx(tmpPath);
  assert(loaded.semantic.sheets[0].comments!.length >= 1, 'round-trip 保留批注');

  // validate
  const issues = validateXlsx(wb);
  assertEqual(issues.filter(i => i.level === 'error').length, 0, '带批注的工作簿校验通过');
}

// ============================================================
// 12. PPTX 批注
// ============================================================

async function testPptxComments() {
  console.log('\n--- PPTX Comments ---');

  const pres = createPptx({ title: '批注测试' });

  // addComment
  const c1 = addPptxComment(pres, 0, '张三', '这里需要修改', 100, 200);
  assert(c1.id !== undefined, 'addComment 返回带 id');
  assertEqual(c1.authorName, '张三', 'authorName 正确');
  assertEqual(c1.text, '这里需要修改', 'text 正确');
  assert(c1.position?.x === 100, 'position.x 正确');

  const c2 = addPptxComment(pres, 0, '李四', '标题太大');
  assertEqual(pres.slides[0].comments!.length, 2, 'slide 有 2 条批注');

  // getCommentText
  assertEqual(getPptxCommentText(pres, 0, c1.id), '这里需要修改', 'getCommentText 正确');

  // listComments
  const all = listPptxComments(pres);
  assertEqual(all.length, 2, 'listComments 返回 2 条');

  // listSlideComments
  const slide0 = listSlideComments(pres, 0);
  assertEqual(slide0.length, 2, 'listSlideComments 返回 2 条');

  // removeComment
  assert(removePptxComment(pres, 0, c2.id), 'removeComment 成功');
  assertEqual(pres.slides[0].comments!.length, 1, '移除后 1 条');

  // round-trip
  const tmpPath = join(OUTPUT_DIR, 'test-pptx-comments.pptx');
  await savePptx(pres, tmpPath);
  const loaded = await loadPptx(tmpPath);
  assert(loaded.semantic.slides[0].comments!.length >= 1, 'round-trip 保留批注');
  assertEqual(loaded.semantic.slides[0].comments![0].authorName, '张三', 'round-trip 保留作者');

  // validate
  const issues = validatePptx(pres);
  assertEqual(issues.filter(i => i.level === 'error').length, 0, '带批注的演示文稿校验通过');
}

// ============================================================
// 主入口
// ============================================================

async function main() {
  console.log('='.repeat(60));
  console.log('New APIs Test');
  console.log('='.repeat(60));
  console.log(`Output: ${OUTPUT_DIR}`);
  console.log(`Time: ${new Date().toISOString()}`);

  await testCreateDocx();
  await testCreateXlsx();
  await testCreatePptx();
  await testUpdateDocxMeta();
  await testUpdateXlsxMeta();
  await testUpdatePptxMeta();
  await testDetectFormat();
  await testValidateBuffer();
  await testToBuffer();
  await testToJSON();
  await testErrorTypes();
  await testFullWorkflow();
  await testComments();
  await testRevisions();
  await testXlsxComments();
  await testPptxComments();

  console.log('\n' + '='.repeat(60));
  console.log(`Results: ${passed} passed, ${failed} failed, ${passed + failed} total`);
  console.log('='.repeat(60));

  if (failed > 0) process.exit(1);
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
