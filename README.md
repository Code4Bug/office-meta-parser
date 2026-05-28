# office-meta-parser

纯 TypeScript 实现的 Office Open XML (OOXML) 解析与序列化库，支持 DOCX、XLSX、PPTX 格式。

零原生依赖，可在 Node.js 和浏览器环境中运行。

## 安装

```bash
npm install office-meta-parser
```

## 快速上手

### 从文件加载

```typescript
import { loadDocx, saveDocx } from 'office-meta-parser/docx';

// 加载 → 修改 → 保存
const { semantic } = await loadDocx('report.docx');
semantic.body.blocks.push({
  type: 'paragraph',
  runs: [{ text: '新增段落', bold: true }],
});
await saveDocx(semantic, 'output.docx');
```

### 从零创建

```typescript
import { createDocx, docx, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: '月度报告', creator: '张三' });

docx.updateTitle(doc, '2024年5月月度报告');
docx.updateCategory(doc, '工作报告');

doc.body.blocks.push(
  { type: 'paragraph', runs: [{ text: '一、概述', bold: true, fontSize: 28 }] },
  { type: 'paragraph', runs: [{ text: '本月完成了核心功能开发。' }] },
);

await saveDocx(doc, 'report.docx');
```

### Buffer 级操作

```typescript
import { parseDocx, serializeDocx } from 'office-meta-parser/docx';

// 解析 ArrayBuffer
const { raw, semantic } = await parseDocx(arrayBuffer);

// 序列化为 ArrayBuffer
const output = await serializeDocx(semantic);
```

---

## 使用指南

### 一、文件读写

#### 加载本地文件

```typescript
import { loadDocx, saveDocx } from 'office-meta-parser/docx';
import { loadXlsx, saveXlsx } from 'office-meta-parser/xlsx';
import { loadPptx, savePptx } from 'office-meta-parser/pptx';

const { semantic: doc } = await loadDocx('input.docx');
const { semantic: wb }  = await loadXlsx('input.xlsx');
const { semantic: pres } = await loadPptx('input.pptx');
```

#### 保存到文件

```typescript
await saveDocx(doc, 'output.docx');
await saveXlsx(wb, 'output.xlsx');
await savePptx(pres, 'output.pptx');
```

#### 写到流（HTTP 响应 / 文件流）

```typescript
import { writeDocxToStream } from 'office-meta-parser/docx';

// Express 示例
app.get('/export', async (req, res) => {
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="report.docx"');
  await writeDocxToStream(doc, res);
});
```

#### Buffer 与 JSON 转换

```typescript
import { toBuffer, toJSON, toJSONString, saveToJSON } from 'office-meta-parser';

const buf = toBuffer(arrayBuffer);     // → Node.js Buffer
const json = toJSON(doc);              // → 可序列化对象（剥离 rawXmlParts）
const str = toJSONString(doc, 2);      // → JSON 字符串
await saveToJSON(doc, 'output.json');   // → 写 JSON 文件
```

---

### 二、DOCX 操作

#### 创建文档并添加内容

```typescript
import { createDocx, docx, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: '项目报告', creator: '张三' });

// 添加标题段落
doc.body.blocks.push({
  type: 'paragraph',
  runs: [{ text: '项目进展报告', bold: true, fontSize: 36, color: '1F4E79' }],
});

// 添加正文段落
doc.body.blocks.push({
  type: 'paragraph',
  runs: [{ text: '本季度完成了以下工作：' }],
});

// 添加带混合格式的段落
doc.body.blocks.push({
  type: 'paragraph',
  runs: [
    { text: '核心模块', bold: true },
    { text: '已通过全部 ' },
    { text: '454', bold: true, color: 'FF0000' },
    { text: ' 个测试用例。' },
  ],
});

await saveDocx(doc, 'report.docx');
```

#### 添加表格

```typescript
doc.body.blocks.push({
  type: 'table',
  rows: [
    {
      cells: [
        { blocks: [{ type: 'paragraph', runs: [{ text: '姓名', bold: true }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: '部门', bold: true }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: '绩效', bold: true }] }] },
      ],
    },
    {
      cells: [
        { blocks: [{ type: 'paragraph', runs: [{ text: '张三' }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: '研发部' }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: 'A' }] }] },
      ],
    },
  ],
});
```

#### 添加超链接

```typescript
doc.body.blocks.push({
  type: 'hyperlink',
  relationshipId: 'rId10',
  url: 'https://example.com',
  runs: [{ text: '访问官网', underline: true, color: '0563C1' }],
});
```

#### 读取文档内容

```typescript
const { semantic } = await loadDocx('input.docx');

// 遍历所有段落
for (const block of semantic.body.blocks) {
  if (block.type === 'paragraph') {
    const text = block.runs.map(r => r.text).join('');
    console.log(text);
  } else if (block.type === 'table') {
    for (const row of block.rows) {
      const cells = row.cells.map(c =>
        c.blocks[0]?.runs?.map(r => r.text).join('') ?? ''
      );
      console.log(cells.join(' | '));
    }
  }
}
```

#### 修改元数据

```typescript
import { docx } from 'office-meta-parser/docx';

const { semantic } = await loadDocx('input.docx');

docx.updateTitle(semantic, '新标题');
docx.updateCreator(semantic, '新作者');
docx.updateCategory(semantic, '合同');
docx.updateLastModifiedBy(semantic, '系统');

await saveDocx(semantic, 'output.docx');
```

#### 批注操作

```typescript
import { createDocx, addComment, listComments, markCommentDone, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: '审阅文档' });
const run = { text: '待审核内容' };
doc.body.blocks.push({ type: 'paragraph', runs: [run] });

// 添加批注
const comment = addComment(doc, run, '审核人', '请补充数据来源');

// 查看批注
const comments = listComments(doc);
console.log(comments[0].comment.content);  // 批注内容

// 标记已完成
markCommentDone(doc, comment.id);

await saveDocx(doc, 'reviewed.docx');
```

#### 修订操作

```typescript
import { createDocx, markInsert, markDelete, hasPendingRevisions, acceptAllInserts, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: '修订文档' });
const run1 = { text: '原始内容' };
const run2 = { text: '新增内容' };
doc.body.blocks.push({ type: 'paragraph', runs: [run1, run2] });

// 标记修订
markInsert(run2, '编辑');
markDelete(run1, '编辑');

// 检查是否有未处理的修订
console.log(hasPendingRevisions(doc));  // true

// 接受所有插入修订
const count = acceptAllInserts(doc);

await saveDocx(doc, 'revised.docx');
```

---

### 三、XLSX 操作

#### 创建工作簿

```typescript
import { createXlsx, saveXlsx } from 'office-meta-parser/xlsx';

const wb = createXlsx({ title: '销售报表', creator: '王五', sheetName: '月度数据' });

// 添加表头
wb.sheets[0].cells.push([
  { value: '产品名称', type: 'string' },
  { value: '销量', type: 'string' },
  { value: '单价', type: 'string' },
  { value: '总额', type: 'string' },
]);

// 添加数据行
wb.sheets[0].cells.push([
  { value: '笔记本电脑', type: 'string' },
  { value: 120, type: 'number' },
  { value: 5999, type: 'number' },
  { value: null, type: 'formula', formula: 'B2*C2' },
]);

// 添加汇总行
wb.sheets[0].cells.push([
  { value: '合计', type: 'string' },
  { value: null, type: 'formula', formula: 'SUM(B2:B3)' },
  { value: null, type: 'string' },
  { value: null, type: 'formula', formula: 'SUM(D2:D3)' },
]);

await saveXlsx(wb, 'sales.xlsx');
```

#### 读取单元格数据

```typescript
const { semantic } = await loadXlsx('input.xlsx');

for (const sheet of semantic.sheets) {
  console.log(`Sheet: ${sheet.name}`);
  for (const row of sheet.cells) {
    for (const cell of row) {
      if (cell.type === 'sharedString') {
        // 共享字符串：通过索引查表
        const text = semantic.sharedStrings[Number(cell.value)];
        console.log(text);
      } else {
        console.log(cell.value);
      }
    }
  }
}
```

#### 添加多个工作表

```typescript
const wb = createXlsx({ title: '年度报表' });

wb.sheets.push(
  { name: 'Q1', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
  { name: 'Q2', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
  { name: 'Q3', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
  { name: 'Q4', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
);
```

#### 单元格批注

```typescript
import { createXlsx, addComment, listComments, saveXlsx } from 'office-meta-parser/xlsx';

const wb = createXlsx({ title: '审核表' });
// 确保 sheet 有单元格数据
wb.sheets[0].cells = [[{ value: '数据', type: 'string', formula: undefined }]];

// 添加批注
addComment(wb, 0, 'A1', '审核人', '请核实数据来源');

// 查看所有批注
console.log(listComments(wb));

await saveXlsx(wb, 'reviewed.xlsx');
```

---

### 四、PPTX 操作

#### 创建演示文稿

```typescript
import { createPptx, savePptx } from 'office-meta-parser/pptx';

const pres = createPptx({ title: '产品介绍', creator: '赵六' });

// 第一页：标题页
pres.slides[0].elements.push(
  {
    type: 'text',
    content: '产品介绍演示文稿',
    position: { x: 0, y: 0, width: 9144000, height: 2000000 },
    paragraphs: [{ runs: [{ text: '产品介绍演示文稿' }] }],
    placeholder: { type: 'title' },
  },
  {
    type: 'text',
    content: '2024年度产品线全面介绍',
    position: { x: 0, y: 3000000, width: 9144000, height: 1000000 },
    paragraphs: [{ runs: [{ text: '2024年度产品线全面介绍' }] }],
    placeholder: { type: 'subtitle' },
  },
);

// 添加新幻灯片
pres.slides.push({
  elements: [
    {
      type: 'text',
      content: '核心产品线',
      position: { x: 0, y: 0, width: 9144000, height: 1000000 },
      paragraphs: [{ runs: [{ text: '核心产品线', bold: true }] }],
    },
    {
      type: 'text',
      content: '笔记本电脑系列\n显示器系列\n外设配件系列',
      position: { x: 0, y: 1500000, width: 9144000, height: 4000000 },
      paragraphs: [
        { runs: [{ text: '笔记本电脑系列' }] },
        { runs: [{ text: '显示器系列' }] },
        { runs: [{ text: '外设配件系列' }] },
      ],
    },
  ],
});

await savePptx(pres, 'intro.pptx');
```

#### 读取幻灯片内容

```typescript
const { semantic } = await loadPptx('input.pptx');

for (let i = 0; i < semantic.slides.length; i++) {
  const slide = semantic.slides[i];
  console.log(`\n--- Slide ${i + 1} ---`);
  for (const el of slide.elements) {
    if (el.type === 'text') {
      console.log(el.content);
    } else if (el.type === 'image') {
      console.log(`[图片: ${el.relationshipId}]`);
    } else if (el.type === 'table') {
      console.log(`[表格: ${el.rows.length} 行]`);
    }
  }
}
```

#### 幻灯片批注

```typescript
import { createPptx, addComment, listComments, savePptx } from 'office-meta-parser/pptx';

const pres = createPptx({ title: '审阅演示' });

// 添加批注（可指定位置坐标）
addComment(pres, 0, '审核人', '标题字号太大', 100, 50);
addComment(pres, 0, '经理', '需要补充数据');

// 查看批注
const comments = listComments(pres);
console.log(comments.length);  // 2

await savePptx(pres, 'reviewed.pptx');
```

---

### 五、格式检测与校验

#### 自动检测文件格式

```typescript
import { detectFormat } from 'office-meta-parser';
import { loadFromFile } from 'office-meta-parser/core';

const buffer = await loadFromFile('unknown.file');
const format = await detectFormat(buffer);

if (format === 'docx') {
  const { semantic } = await parseDocx(buffer);
  // ...
} else if (format === 'xlsx') {
  const { semantic } = await parseXlsx(buffer);
  // ...
} else if (format === 'pptx') {
  const { semantic } = await parsePptx(buffer);
  // ...
} else {
  console.error('不支持的文件格式');
}
```

#### 上传文件校验

```typescript
import { validate } from 'office-meta-parser';

app.post('/upload', async (req, res) => {
  const buffer = req.file.buffer;
  const result = await validate(buffer);

  if (!result.valid) {
    return res.status(400).json({
      error: '文件校验失败',
      format: result.format,
      issues: result.issues.filter(i => i.level === 'error'),
    });
  }

  // 继续处理...
});
```

#### 序列化前自动校验

`serializeDocx` / `serializeXlsx` / `serializePptx` 内部自动执行校验，遇到 error 级别问题会抛出 `ValidationError`：

```typescript
import { ValidationError } from 'office-meta-parser';

try {
  await saveDocx(doc, 'output.docx');
} catch (e) {
  if (e instanceof ValidationError) {
    console.error('校验失败:');
    for (const issue of e.issues) {
      console.error(`  [${issue.level}] ${issue.path}: ${issue.message}`);
    }
  }
}
```

---

### 六、批量处理

#### 批量修改元数据

```typescript
import { loadDocx, saveDocx, docx } from 'office-meta-parser/docx';
import { readdir } from 'fs/promises';

const files = (await readdir('./contracts')).filter(f => f.endsWith('.docx'));

for (const file of files) {
  const { semantic } = await loadDocx(`./contracts/${file}`);
  docx.updateCategory(semantic, '合同');
  docx.updateLastModifiedBy(semantic, '批量归档系统');
  await saveDocx(semantic, `./output/${file}`);
  console.log(`已处理: ${file}`);
}
```

#### 批量导出为 JSON

```typescript
import { loadXlsx, xlsx } from 'office-meta-parser/xlsx';

const files = ['report-q1.xlsx', 'report-q2.xlsx', 'report-q3.xlsx'];

for (const file of files) {
  const { semantic } = await loadXlsx(`./data/${file}`);
  const jsonName = file.replace('.xlsx', '.json');
  await xlsx.saveJSON(semantic, `./json/${jsonName}`);
}
```

---

### 七、Express / Koa 集成

```typescript
import express from 'express';
import { createDocx, docx, writeDocxToStream } from 'office-meta-parser/docx';
import { createXlsx, writeXlsxToStream } from 'office-meta-parser/xlsx';
import { createPptx, writePptxToStream } from 'office-meta-parser/pptx';

const app = express();

// 导出 DOCX
app.get('/api/export/docx', async (req, res) => {
  const doc = createDocx({ title: '导出报告' });
  doc.body.blocks.push({ type: 'paragraph', runs: [{ text: '动态生成的内容' }] });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="report.docx"');
  await writeDocxToStream(doc, res);
});

// 导出 XLSX
app.get('/api/export/xlsx', async (req, res) => {
  const wb = createXlsx({ title: '数据导出', sheetName: 'Sheet1' });
  wb.sheets[0].cells.push([
    { value: 'ID', type: 'string' },
    { value: '名称', type: 'string' },
  ]);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="data.xlsx"');
  await writeXlsxToStream(wb, res);
});

// 导出 PPTX
app.get('/api/export/pptx', async (req, res) => {
  const pres = createPptx({ title: '自动生成' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', 'attachment; filename="slides.pptx"');
  await writePptxToStream(pres, res);
});
```

---

## API 参考

### 导入路径

| 路径 | 内容 |
|------|------|
| `office-meta-parser` | 核心层（detectFormat / validate / toBuffer / toJSON 等） |
| `office-meta-parser/core` | 基础设施（XML / ZIP / 元数据） |
| `office-meta-parser/docx` | Word 文档 |
| `office-meta-parser/xlsx` | Excel 表格 |
| `office-meta-parser/pptx` | PowerPoint 演示文稿 |

---

### DOCX (`office-meta-parser/docx`)

#### 解析 / 序列化

| 函数 | 签名 | 说明 |
|------|------|------|
| `parseDocx` | `(buffer: ArrayBuffer) → Promise<{raw, semantic}>` | 解码 |
| `serializeDocx` | `(doc: DocxDocument) → Promise<ArrayBuffer>` | 编码（含校验） |
| `loadDocx` | `(path: string) → Promise<{raw, semantic}>` | 从文件加载 |
| `saveDocx` | `(doc: DocxDocument, path: string) → Promise<void>` | 保存到文件 |
| `writeDocxToStream` | `(doc, stream: Writable) → Promise<void>` | 写到流 |
| `createDocx` | `(options?) → DocxDocument` | 创建空文档 |
| `validateDocx` | `(doc: DocxDocument) → ValidationIssue[]` | 校验语义模型 |

#### 元数据更新

```typescript
import { docx, updateDocxTitle, updateDocxCreator } from 'office-meta-parser/docx';

// 方式 1: 命名空间
docx.updateTitle(doc, '新标题');
docx.updateCreator(doc, '张三');
docx.updateSubject(doc, '主题');
docx.updateDescription(doc, '描述');
docx.updateKeywords(doc, '关键词');
docx.updateCategory(doc, '分类');
docx.updateLastModifiedBy(doc, '李四');

// 方式 2: 独立函数
updateDocxTitle(doc, '新标题');
```

#### JSON 导出

```typescript
import { toDocxJSON, toDocxJSONString, saveDocxJSON, docx } from 'office-meta-parser/docx';

const json = toDocxJSON(doc);              // → DocxDocument 对象
const str  = toDocxJSONString(doc, 2);     // → JSON 字符串
await saveDocxJSON(doc, 'output.json');     // → 写文件

// 命名空间方式
docx.toJSON(doc);
docx.toJSONString(doc);
docx.saveJSON(doc, 'output.json');
```

#### 批注操作

```typescript
import { addComment, removeComment, listComments, getCommentText, markCommentDone, markCommentUndone } from 'office-meta-parser/docx';

// 添加批注 — 将批注锚定到指定 TextRun
const comment = addComment(doc, run, '张三', '这里需要修改');
// run.commentId 已自动设置

// 列出所有批注
const comments = listComments(doc);
// [{ comment, blockIndex, runIndices }, ...]

// 获取批注纯文本
const text = getCommentText(doc, comment.id);  // → '这里需要修改'

// 标记完成 / 未完成
markCommentDone(doc, comment.id);
markCommentUndone(doc, comment.id);

// 移除批注（同时清除 run 上的 commentId 标记）
removeComment(doc, comment.id);
```

#### 修订操作

```typescript
import {
  markInsert, markDelete, addFormatChange, clearRevision,
  listRevisions, hasPendingRevisions,
  acceptAllInserts, acceptAllDeletes, rejectAllInserts, rejectAllDeletes,
} from 'office-meta-parser/docx';

// 标记插入 / 删除修订
markInsert(run, '张三', '2024-01-01T00:00:00Z');
markDelete(run, '张三');

// 添加格式变更记录
addFormatChange(doc, '李四');

// 查询修订
hasPendingRevisions(doc);          // → boolean
const revisions = listRevisions(doc);  // [{ revision, blockIndex, runIndex, run }, ...]

// 接受 / 拒绝
acceptAllInserts(doc);   // → 处理数量（移除标记，保留文本）
acceptAllDeletes(doc);   // → 处理数量（移除被标记的 run）
rejectAllInserts(doc);   // → 处理数量（移除被标记的 run）
rejectAllDeletes(doc);   // → 处理数量（移除标记，保留文本）

// 清除单个 run 的修订标记
clearRevision(run);
```

---

### XLSX (`office-meta-parser/xlsx`)

#### 解析 / 序列化

| 函数 | 签名 | 说明 |
|------|------|------|
| `parseXlsx` | `(buffer: ArrayBuffer) → Promise<{raw, semantic}>` | 解码 |
| `serializeXlsx` | `(wb: XlsxWorkbook) → Promise<ArrayBuffer>` | 编码（含校验） |
| `loadXlsx` | `(path: string) → Promise<{raw, semantic}>` | 从文件加载 |
| `saveXlsx` | `(wb: XlsxWorkbook, path: string) → Promise<void>` | 保存到文件 |
| `writeXlsxToStream` | `(wb, stream: Writable) → Promise<void>` | 写到流 |
| `createXlsx` | `(options?) → XlsxWorkbook` | 创建空工作簿 |
| `validateXlsx` | `(wb: XlsxWorkbook) → ValidationIssue[]` | 校验 |

#### 元数据更新

```typescript
import { xlsx } from 'office-meta-parser/xlsx';

xlsx.updateTitle(wb, '销售报表');
xlsx.updateCreator(wb, '王五');
```

#### JSON 导出

```typescript
import { toXlsxJSON, saveXlsxJSON, xlsx } from 'office-meta-parser/xlsx';

xlsx.toJSON(wb);
xlsx.saveJSON(wb, 'workbook.json');
```

#### 批注操作

```typescript
import { addComment, removeComment, listComments, getCommentText, updateComment } from 'office-meta-parser/xlsx';

// 添加批注 — 按单元格引用定位
addComment(wb, 0, 'A1', '张三', '需要修改');
addComment(wb, 0, 'B2', '李四', '数据有误', [{ text: '富文本', bold: true }]);

// 列出批注
const all = listComments(wb);               // 跨所有 sheet
const sheet0 = listSheetComments(wb, 0);    // 指定 sheet

// 获取 / 更新
getCommentText(wb, 0, 'A1');                // → '需要修改'
updateComment(wb, 0, 'A1', '已修改');

// 移除
removeComment(wb, 0, 'B2');
```

---

### PPTX (`office-meta-parser/pptx`)

#### 解析 / 序列化

| 函数 | 签名 | 说明 |
|------|------|------|
| `parsePptx` | `(buffer: ArrayBuffer) → Promise<{raw, semantic}>` | 解码 |
| `serializePptx` | `(pres: PptxPresentation) → Promise<ArrayBuffer>` | 编码（含校验） |
| `loadPptx` | `(path: string) → Promise<{raw, semantic}>` | 从文件加载 |
| `savePptx` | `(pres: PptxPresentation, path: string) → Promise<void>` | 保存到文件 |
| `writePptxToStream` | `(pres, stream: Writable) → Promise<void>` | 写到流 |
| `createPptx` | `(options?) → PptxPresentation` | 创建空演示文稿 |
| `validatePptx` | `(pres: PptxPresentation) → ValidationIssue[]` | 校验 |

#### 元数据更新

```typescript
import { pptx } from 'office-meta-parser/pptx';

pptx.updateTitle(pres, '产品介绍');
pptx.updateCreator(pres, '赵六');
```

#### JSON 导出

```typescript
import { toPptxJSON, savePptxJSON, pptx } from 'office-meta-parser/pptx';

pptx.toJSON(pres);
pptx.saveJSON(pres, 'presentation.json');
```

#### 批注操作

```typescript
import { addComment, removeComment, listComments, listSlideComments, getCommentText } from 'office-meta-parser/pptx';

// 添加批注 — 可指定位置坐标
addComment(pres, 0, '审核人', '标题需要修改', 100, 200);

// 列出批注
const all = listComments(pres);               // 跨所有 slide
const slide0 = listSlideComments(pres, 0);    // 指定 slide

// 获取文本
getCommentText(pres, 0, '1');                 // → '标题需要修改'

// 移除
removeComment(pres, 0, '1');
```

---

### 核心模块 (`office-meta-parser/core`)

#### 文件 I/O

| 函数 | 签名 | 说明 |
|------|------|------|
| `loadFromFile` | `(path: string) → Promise<ArrayBuffer>` | 读文件 |
| `saveToFile` | `(buffer: ArrayBuffer, path: string) → Promise<void>` | 写文件 |
| `writeToStream` | `(buffer, stream: Writable) → Promise<void>` | 写到流 |
| `toBuffer` | `(buffer: ArrayBuffer) → Buffer` | 转 Node.js Buffer |
| `toJSON` | `<T>(semantic: T) → T` | 语义模型 → JSON 对象 |
| `toJSONString` | `<T>(semantic: T, space?) → string` | 语义模型 → JSON 字符串 |
| `saveToJSON` | `<T>(semantic: T, path: string, space?) → Promise<void>` | 语义模型 → JSON 文件 |

#### 格式检测与校验

```typescript
import { detectFormat, validate } from 'office-meta-parser';

// 检测格式
const format = await detectFormat(buffer); // 'docx' | 'xlsx' | 'pptx' | null

// 统一校验（自动检测格式 + 解析 + 校验）
const result = await validate(buffer);
// { format: 'docx', issues: [...], valid: true }
```

#### 元数据更新（泛型）

```typescript
import { updateTitle, updateCreator, updateSubject, updateDescription,
         updateKeywords, updateCategory, updateLastModifiedBy } from 'office-meta-parser';

// 适用于任意格式的语义模型
updateTitle(doc, '新标题');
updateCreator(wb, '张三');
```

#### XML / ZIP

| 函数 | 说明 |
|------|------|
| `parseXml(xml)` | XML 字符串 → ParsedNode 树 |
| `serializeXml(node)` | ParsedNode 树 → XML 字符串 |
| `parseRels(xml)` | 解析 .rels 关系文件 |
| `serializeRels(rels)` | 序列化 .rels |
| `parseContentTypes(xml)` | 解析 [Content_Types].xml |
| `serializeContentTypes(cts)` | 序列化 Content_Types |
| `unzip(buffer)` | 解压 ZIP → ZipEntry[] |
| `zip(entries)` | 打包为 ZIP |

#### 元数据解析 / 序列化

| 函数 | 说明 |
|------|------|
| `parseMeta(node)` | 解析 core.xml 元数据 |
| `serializeMeta(meta)` | 序列化 core.xml |
| `parseAppMeta(xml)` | 解析 app.xml |
| `serializeAppMeta(meta)` | 序列化 app.xml |
| `parseCustomProperties(xml)` | 解析 custom.xml |
| `serializeCustomProperties(props)` | 序列化 custom.xml |

---

### 错误处理

```typescript
import { ValidationError, FormatError } from 'office-meta-parser';

try {
  await serializeDocx(doc);
} catch (e) {
  if (e instanceof ValidationError) {
    // 校验失败：e.issues 包含具体的错误列表
    for (const issue of e.issues) {
      console.error(`[${issue.level}] ${issue.path}: ${issue.message}`);
    }
  }
}

// FormatError 用于格式不支持的情况
const err = new FormatError('不支持的格式', 'pdf');
```

---

### 校验规则

每个格式的校验器包含**必须项**（error，阻止序列化）和**可选项**（warning，打印到 stderr）。

#### DOCX

| 级别 | 规则 |
|------|------|
| error | body.blocks 不能为空 |
| error | commentId 引用的 comment 必须存在 |
| error | bookmarkStart/bookmarkEnd 必须成对 |
| error | hyperlink/image 必须有 relationshipId |
| error | numbering.numId 引用必须存在 |
| error | header/footer id 不能重复 |
| error | table.rows 不能为空 |
| warning | comment 内容不能为空 |
| warning | 定义了 comments 但 body 中无引用 |

#### XLSX

| 级别 | 规则 |
|------|------|
| error | sheets 不能为空 |
| error | sheet.name 不能为空 |
| error | sharedString 引用不能越界 |
| error | mergedCells 范围合法 |
| error | hyperlinks 必须有 ref 和 url |
| error | table 必须有 ref 和 displayName |

#### PPTX

| 级别 | 规则 |
|------|------|
| error | theme 不能缺失 |
| error | masters/layouts 不能为空 |
| error | slides 不能为空 |
| error | image 必须有 relationshipId |
| error | table.rows 不能为空 |
| warning | layout 引用必须存在 |
| warning | slideSize 缺失 |

---

## 数据模型

### DocumentMeta

```typescript
interface DocumentMeta {
  title?: string;
  subject?: string;
  creator?: string;
  description?: string;
  keywords?: string;
  lastModifiedBy?: string;
  created?: string;
  modified?: string;
  revision?: string;
  category?: string;
}
```

### DOCX

```typescript
interface DocxDocument {
  meta: DocumentMeta;
  styles: StyleDefinitions;
  body: DocxBody;
  comments?: Comment[];
  trackChanges?: Revision[];
  headers?: Header[];
  footers?: Footer[];
  numbering?: NumberingDefinitions;
  footnotes?: Footnote[];
  endnotes?: Footnote[];
}

type DocxBlock = Paragraph | Table | Image | Hyperlink | BookmarkStart | BookmarkEnd;

interface Paragraph {
  type: 'paragraph';
  style?: string;
  runs: TextRun[];
  numbering?: NumberingProperties;
}

interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontSize?: number;       // 半磅值（如 56 = 28pt）
  color?: string;          // RGB 十六进制
  fontFamily?: string;
  highlight?: string;
  commentId?: string;      // 关联批注
}
```

### XLSX

```typescript
interface XlsxWorkbook {
  meta: DocumentMeta;
  sheets: Sheet[];
  styles: CellStyleDefinitions;
  sharedStrings: SharedStringEntry[];
}

interface Sheet {
  name: string;
  cells: Cell[][];
  mergedCells: MergedCell[];
  hyperlinks: Hyperlink[];
  autoFilter?: AutoFilter;
  dataValidations?: DataValidation[];
  conditionalFormats?: ConditionalFormat[];
  frozenPanes?: FrozenPanes;
}

interface Cell {
  type: 'string' | 'sharedString' | 'number' | 'boolean' | 'formula' | 'error';
  value: string | number | boolean | null;
  formula?: string;
}
```

### PPTX

```typescript
interface PptxPresentation {
  meta: DocumentMeta;
  slides: Slide[];
  masters: SlideMaster[];
  layouts: SlideLayout[];
  theme?: Theme;
  slideSize?: { width: number; height: number };
}

interface Slide {
  elements: SlideElement[];
  transition?: Transition;
  notes?: string;
  animations?: Animation[];
}

type SlideElement = TextShape | ImageShape | GroupShape | TableShape | MediaShape;
```

---

## 测试

```bash
npm test              # 运行全部测试（89 文件 / 454 用例）
npm run test:codec    # 编解码集成测试
npm run typecheck     # 类型检查
```

## 依赖

| 包 | 用途 |
|----|------|
| `jszip` | ZIP 压缩/解压 |
| `fast-xml-parser` | XML 解析与序列化 |

## 许可

MIT
