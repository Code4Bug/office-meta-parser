# office-meta-parser

[![npm version](https://img.shields.io/npm/v/@turing-weique/office-meta-parser)](https://www.npmjs.com/package/@turing-weique/office-meta-parser)[![npm downloads](https://img.shields.io/npm/dm/@turing-weique/office-meta-parser)](https://www.npmjs.com/package/@turing-weique/office-meta-parser)[![license](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](./LICENSE)

[中文](./README.md) | English

A pure TypeScript library for parsing and serializing Office Open XML (OOXML) documents — DOCX, XLSX, and PPTX.

Zero native dependencies. Runs in both Node.js and browser environments.

## Installation

```bash
npm install @turing-weique/office-meta-parser
```

## Quick Start

### Unified Namespace (OMP)

```typescript
import { OMP } from '@turing-weique/office-meta-parser';

// Common APIs
const format = await OMP.detectFormat(buffer);
const buf = OMP.toBuffer(arrayBuffer);

// DOCX
const doc = OMP.docx.create({ title: 'Report', creator: 'Alice' });
doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'Content' }] });
OMP.docx.updateTitle(doc, 'New Title');
await OMP.docx.save(doc, 'output.docx');

// XLSX
const wb = OMP.xlsx.create({ title: 'Sales Report' });
OMP.xlsx.addComment(wb, 0, 'A1', 'Reviewer', 'Please verify');

// PPTX
const pres = OMP.pptx.create({ title: 'Slides' });
OMP.pptx.addComment(pres, 0, 'Reviewer', 'Title needs revision');
```

### Load from File

```typescript
import { loadDocx, saveDocx } from 'office-meta-parser/docx';

// Load → Modify → Save
const { semantic } = await loadDocx('report.docx');
semantic.body.blocks.push({
  type: 'paragraph',
  runs: [{ text: 'New paragraph', bold: true }],
});
await saveDocx(semantic, 'output.docx');
```

### Create from Scratch

```typescript
import { createDocx, docx, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: 'Monthly Report', creator: 'Alice' });

docx.updateTitle(doc, 'Monthly Report — May 2024');
docx.updateCategory(doc, 'Work Report');

doc.body.blocks.push(
  { type: 'paragraph', runs: [{ text: '1. Overview', bold: true, fontSize: 28 }] },
  { type: 'paragraph', runs: [{ text: 'Core features were completed this month.' }] },
);

await saveDocx(doc, 'report.docx');
```

### Buffer-Level Operations

```typescript
import { parseDocx, serializeDocx } from 'office-meta-parser/docx';

// Parse ArrayBuffer
const { raw, semantic } = await parseDocx(arrayBuffer);

// Serialize to ArrayBuffer
const output = await serializeDocx(semantic);
```

---

## Usage Guide

### 1. File I/O

#### Load Local Files

```typescript
import { loadDocx, saveDocx } from 'office-meta-parser/docx';
import { loadXlsx, saveXlsx } from 'office-meta-parser/xlsx';
import { loadPptx, savePptx } from 'office-meta-parser/pptx';

const { semantic: doc }  = await loadDocx('input.docx');
const { semantic: wb }   = await loadXlsx('input.xlsx');
const { semantic: pres } = await loadPptx('input.pptx');
```

#### Save to File

```typescript
await saveDocx(doc, 'output.docx');
await saveXlsx(wb, 'output.xlsx');
await savePptx(pres, 'output.pptx');
```

#### Write to Stream (HTTP Response / File Stream)

```typescript
import { writeDocxToStream } from 'office-meta-parser/docx';

// Express example
app.get('/export', async (req, res) => {
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="report.docx"');
  await writeDocxToStream(doc, res);
});
```

#### Buffer & JSON Conversion

```typescript
import { toBuffer, toJSON, toJSONString, saveToJSON } from '@turing-weique/office-meta-parser';

const buf = toBuffer(arrayBuffer);     // → Node.js Buffer
const json = toJSON(doc);              // → Serializable object (strips rawXmlParts)
const str = toJSONString(doc, 2);      // → JSON string
await saveToJSON(doc, 'output.json');   // → Write JSON file
```

---

### 2. DOCX Operations

#### Create a Document and Add Content

```typescript
import { createDocx, docx, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: 'Project Report', creator: 'Alice' });

// Add a heading paragraph
doc.body.blocks.push({
  type: 'paragraph',
  runs: [{ text: 'Project Progress Report', bold: true, fontSize: 36, color: '1F4E79' }],
});

// Add a body paragraph
doc.body.blocks.push({
  type: 'paragraph',
  runs: [{ text: 'The following work was completed this quarter:' }],
});

// Add a paragraph with mixed formatting
doc.body.blocks.push({
  type: 'paragraph',
  runs: [
    { text: 'Core modules', bold: true },
    { text: ' passed all ' },
    { text: '454', bold: true, color: 'FF0000' },
    { text: ' test cases.' },
  ],
});

await saveDocx(doc, 'report.docx');
```

#### Add a Table

```typescript
doc.body.blocks.push({
  type: 'table',
  rows: [
    {
      cells: [
        { blocks: [{ type: 'paragraph', runs: [{ text: 'Name', bold: true }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: 'Department', bold: true }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: 'Rating', bold: true }] }] },
      ],
    },
    {
      cells: [
        { blocks: [{ type: 'paragraph', runs: [{ text: 'Alice' }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: 'Engineering' }] }] },
        { blocks: [{ type: 'paragraph', runs: [{ text: 'A' }] }] },
      ],
    },
  ],
});
```

#### Add a Hyperlink

```typescript
doc.body.blocks.push({
  type: 'hyperlink',
  relationshipId: 'rId10',
  url: 'https://example.com',
  runs: [{ text: 'Visit website', underline: true, color: '0563C1' }],
});
```

#### Read Document Content

```typescript
const { semantic } = await loadDocx('input.docx');

// Iterate over all blocks
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

#### Update Metadata

```typescript
import { OMP } from '@turing-weique/office-meta-parser';

const { semantic } = await OMP.docx.load('input.docx');

OMP.docx.updateTitle(semantic, 'New Title');
OMP.docx.updateCreator(semantic, 'New Author');
OMP.docx.updateCategory(semantic, 'Contract');
OMP.docx.updateLastModifiedBy(semantic, 'System');

await OMP.docx.save(semantic, 'output.docx');
```

#### Comment Operations

```typescript
import { createDocx, addComment, listComments, getCommentText, markCommentDone, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: 'Review Document' });
const run = { text: 'Content under review' };
doc.body.blocks.push({ type: 'paragraph', runs: [run] });

// Add a comment anchored to a TextRun
const comment = addComment(doc, run, 'Reviewer', 'Please add data sources');

// List all comments
const comments = listComments(doc);
console.log(comments.length);  // 1

// Get comment plain text
console.log(getCommentText(doc, comment.id));  // → 'Please add data sources'

// Mark as done
markCommentDone(doc, comment.id);

await saveDocx(doc, 'reviewed.docx');
```

#### Revision (Track Changes) Operations

```typescript
import { createDocx, markInsert, markDelete, hasPendingRevisions, acceptAllInserts, saveDocx } from 'office-meta-parser/docx';

const doc = createDocx({ title: 'Revised Document' });
const run1 = { text: 'Original content' };
const run2 = { text: 'New content' };
doc.body.blocks.push({ type: 'paragraph', runs: [run1, run2] });

// Mark revisions
markInsert(run2, 'Editor');
markDelete(run1, 'Editor');

// Check for pending revisions
console.log(hasPendingRevisions(doc));  // true

// Accept all insertions
const count = acceptAllInserts(doc);

await saveDocx(doc, 'revised.docx');
```

---

### 3. XLSX Operations

#### Create a Workbook

```typescript
import { createXlsx, saveXlsx } from 'office-meta-parser/xlsx';

const wb = createXlsx({ title: 'Sales Report', creator: 'Bob', sheetName: 'Monthly Data' });

// Add headers
wb.sheets[0].cells.push([
  { value: 'Product', type: 'string' },
  { value: 'Quantity', type: 'string' },
  { value: 'Unit Price', type: 'string' },
  { value: 'Total', type: 'string' },
]);

// Add data rows
wb.sheets[0].cells.push([
  { value: 'Laptop', type: 'string' },
  { value: 120, type: 'number' },
  { value: 999, type: 'number' },
  { value: null, type: 'formula', formula: 'B2*C2' },
]);

// Add a summary row
wb.sheets[0].cells.push([
  { value: 'Total', type: 'string' },
  { value: null, type: 'formula', formula: 'SUM(B2:B3)' },
  { value: null, type: 'string' },
  { value: null, type: 'formula', formula: 'SUM(D2:D3)' },
]);

await saveXlsx(wb, 'sales.xlsx');
```

#### Read Cell Data

```typescript
const { semantic } = await loadXlsx('input.xlsx');

for (const sheet of semantic.sheets) {
  console.log(`Sheet: ${sheet.name}`);
  for (const row of sheet.cells) {
    for (const cell of row) {
      if (cell.type === 'sharedString') {
        // Shared string: look up by index
        const text = semantic.sharedStrings[Number(cell.value)];
        console.log(text);
      } else {
        console.log(cell.value);
      }
    }
  }
}
```

#### Add Multiple Sheets

```typescript
const wb = createXlsx({ title: 'Annual Report' });

wb.sheets.push(
  { name: 'Q1', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
  { name: 'Q2', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
  { name: 'Q3', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
  { name: 'Q4', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] },
);
```

#### Cell Comments

```typescript
import { createXlsx, addComment, listComments, listSheetComments, getCommentText, saveXlsx } from 'office-meta-parser/xlsx';

const wb = createXlsx({ title: 'Review Sheet' });
// Ensure the sheet has cell data
wb.sheets[0].cells = [[{ value: 'Data', type: 'string' }]];

// Add a comment
addComment(wb, 0, 'A1', 'Reviewer', 'Please verify the data source');

// List comments
console.log(listComments(wb));              // Across all sheets
console.log(listSheetComments(wb, 0));      // Specific sheet
console.log(getCommentText(wb, 0, 'A1'));   // → 'Please verify the data source'

await saveXlsx(wb, 'reviewed.xlsx');
```

---

### 4. PPTX Operations

#### Create a Presentation

```typescript
import { createPptx, savePptx } from 'office-meta-parser/pptx';

const pres = createPptx({ title: 'Product Intro', creator: 'Charlie' });

// Slide 1: Title slide
pres.slides[0].elements.push(
  {
    type: 'text',
    content: 'Product Introduction',
    position: { x: 0, y: 0, width: 9144000, height: 2000000 },
    paragraphs: [{ runs: [{ text: 'Product Introduction' }] }],
    placeholder: { type: 'title' },
  },
  {
    type: 'text',
    content: 'Full product line overview for 2024',
    position: { x: 0, y: 3000000, width: 9144000, height: 1000000 },
    paragraphs: [{ runs: [{ text: 'Full product line overview for 2024' }] }],
    placeholder: { type: 'subtitle' },
  },
);

// Add a new slide
pres.slides.push({
  elements: [
    {
      type: 'text',
      content: 'Core Products',
      position: { x: 0, y: 0, width: 9144000, height: 1000000 },
      paragraphs: [{ runs: [{ text: 'Core Products', bold: true }] }],
    },
    {
      type: 'text',
      content: 'Laptop Series\nMonitor Series\nPeripheral Series',
      position: { x: 0, y: 1500000, width: 9144000, height: 4000000 },
      paragraphs: [
        { runs: [{ text: 'Laptop Series' }] },
        { runs: [{ text: 'Monitor Series' }] },
        { runs: [{ text: 'Peripheral Series' }] },
      ],
    },
  ],
});

await savePptx(pres, 'intro.pptx');
```

#### Read Slide Content

```typescript
const { semantic } = await loadPptx('input.pptx');

for (let i = 0; i < semantic.slides.length; i++) {
  const slide = semantic.slides[i];
  console.log(`\n--- Slide ${i + 1} ---`);
  for (const el of slide.elements) {
    if (el.type === 'text') {
      console.log(el.content);
    } else if (el.type === 'image') {
      console.log(`[Image: ${el.relationshipId}]`);
    } else if (el.type === 'table') {
      console.log(`[Table: ${el.rows.length} rows]`);
    }
  }
}
```

#### Slide Comments

```typescript
import { createPptx, addComment, listComments, listSlideComments, savePptx } from 'office-meta-parser/pptx';

const pres = createPptx({ title: 'Review Presentation' });

// Add comments (with optional position coordinates)
addComment(pres, 0, 'Reviewer', 'Title font is too large', 100, 50);
addComment(pres, 0, 'Manager', 'Need more data');

// List comments
console.log(listComments(pres));            // Across all slides
console.log(listSlideComments(pres, 0));    // Specific slide

await savePptx(pres, 'reviewed.pptx');
```

---

### 5. Format Detection & Validation

#### Auto-Detect File Format

```typescript
import { detectFormat } from '@turing-weique/office-meta-parser';
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
  console.error('Unsupported file format');
}
```

#### Upload Validation

```typescript
import { validate } from '@turing-weique/office-meta-parser';

app.post('/upload', async (req, res) => {
  const buffer = req.file.buffer;
  const result = await validate(buffer);

  if (!result.valid) {
    return res.status(400).json({
      error: 'Validation failed',
      format: result.format,
      issues: result.issues.filter(i => i.level === 'error'),
    });
  }

  // Continue processing...
});
```

#### Auto-Validation on Serialize

`serializeDocx` / `serializeXlsx` / `serializePptx` automatically run validation internally. Error-level issues throw a `ValidationError`:

```typescript
import { ValidationError } from '@turing-weique/office-meta-parser';

try {
  await saveDocx(doc, 'output.docx');
} catch (e) {
  if (e instanceof ValidationError) {
    console.error('Validation failed:');
    for (const issue of e.issues) {
      console.error(`  [${issue.level}] ${issue.path}: ${issue.message}`);
    }
  }
}
```

---

### 6. Batch Processing

#### Batch Metadata Update

```typescript
import { loadDocx, saveDocx, docx } from 'office-meta-parser/docx';
import { readdir } from 'fs/promises';

const files = (await readdir('./contracts')).filter(f => f.endsWith('.docx'));

for (const file of files) {
  const { semantic } = await loadDocx(`./contracts/${file}`);
  docx.updateCategory(semantic, 'Contract');
  docx.updateLastModifiedBy(semantic, 'Batch Archive System');
  await saveDocx(semantic, `./output/${file}`);
  console.log(`Processed: ${file}`);
}
```

#### Batch Export to JSON

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

### 7. Express / Koa Integration

```typescript
import express from 'express';
import { createDocx, docx, writeDocxToStream } from 'office-meta-parser/docx';
import { createXlsx, writeXlsxToStream } from 'office-meta-parser/xlsx';
import { createPptx, writePptxToStream } from 'office-meta-parser/pptx';

const app = express();

// Export DOCX
app.get('/api/export/docx', async (req, res) => {
  const doc = createDocx({ title: 'Export Report' });
  doc.body.blocks.push({ type: 'paragraph', runs: [{ text: 'Dynamically generated content' }] });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="report.docx"');
  await writeDocxToStream(doc, res);
});

// Export XLSX
app.get('/api/export/xlsx', async (req, res) => {
  const wb = createXlsx({ title: 'Data Export', sheetName: 'Sheet1' });
  wb.sheets[0].cells.push([
    { value: 'ID', type: 'string' },
    { value: 'Name', type: 'string' },
  ]);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="data.xlsx"');
  await writeXlsxToStream(wb, res);
});

// Export PPTX
app.get('/api/export/pptx', async (req, res) => {
  const pres = createPptx({ title: 'Auto-Generated' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', 'attachment; filename="slides.pptx"');
  await writePptxToStream(pres, res);
});
```

---

## API Reference

### Import Paths

| Path | Contents |
|------|----------|
| `@turing-weique/office-meta-parser` | Core + `OMP` unified namespace |
| `@turing-weique/office-meta-parser/core` | Infrastructure (XML / ZIP / metadata) |
| `@turing-weique/office-meta-parser/docx` | Word documents |
| `@turing-weique/office-meta-parser/xlsx` | Excel spreadsheets |
| `@turing-weique/office-meta-parser/pptx` | PowerPoint presentations |

### OMP Unified Namespace

All APIs are accessible through the `OMP` object:

```typescript
import { OMP } from '@turing-weique/office-meta-parser';

// Common
OMP.detectFormat(buffer)       // Detect format
OMP.validate(buffer)           // Validate
OMP.toBuffer(arrayBuffer)      // Convert to Buffer
OMP.toJSON(semantic)           // Convert to JSON
OMP.loadFromFile(path)         // Read file
OMP.saveToFile(buffer, path)   // Write file

// Format-specific
OMP.docx.create / .parse / .serialize / .load / .save / .validate
OMP.xlsx.create / .parse / .serialize / .load / .save / .validate
OMP.pptx.create / .parse / .serialize / .load / .save / .validate

// Metadata
OMP.docx.updateTitle / .updateCreator / .updateSubject / ...
OMP.xlsx.updateTitle / .updateCreator / ...
OMP.pptx.updateTitle / .updateCreator / ...

// Comments
OMP.docx.addComment / .removeComment / .listComments / ...
OMP.xlsx.addComment / .removeComment / .listComments / ...
OMP.pptx.addComment / .removeComment / .listComments / ...

// Revisions (DOCX only)
OMP.docx.markInsert / .markDelete / .acceptAllInserts / ...

// JSON
OMP.docx.toJSON / .toJSONString / .saveJSON
OMP.xlsx.toJSON / .toJSONString / .saveJSON
OMP.pptx.toJSON / .toJSONString / .saveJSON
```

---

### DOCX (`@turing-weique/office-meta-parser/docx`)

#### Parse / Serialize

| Function | Signature | Description |
|----------|-----------|-------------|
| `parseDocx` | `(buffer: ArrayBuffer) → Promise<{raw, semantic}>` | Decode |
| `serializeDocx` | `(doc: DocxDocument) → Promise<ArrayBuffer>` | Encode (with validation) |
| `loadDocx` | `(path: string) → Promise<{raw, semantic}>` | Load from file |
| `saveDocx` | `(doc: DocxDocument, path: string) → Promise<void>` | Save to file |
| `writeDocxToStream` | `(doc, stream: Writable) → Promise<void>` | Write to stream |
| `createDocx` | `(options?) → DocxDocument` | Create empty document |
| `validateDocx` | `(doc: DocxDocument) → ValidationIssue[]` | Validate semantic model |

#### Metadata Update

```typescript
import { docx, updateDocxTitle } from 'office-meta-parser/docx';
import { OMP } from '@turing-weique/office-meta-parser';

// Method 1: OMP unified namespace
OMP.docx.updateTitle(doc, 'New Title');
OMP.docx.updateCreator(doc, 'Alice');

// Method 2: Namespace
docx.updateTitle(doc, 'New Title');

// Method 3: Standalone function
updateDocxTitle(doc, 'New Title');

// 7 fields available: updateTitle / updateSubject / updateCreator /
// updateDescription / updateKeywords / updateCategory / updateLastModifiedBy
```

#### JSON Export

```typescript
import { toDocxJSON, docx } from 'office-meta-parser/docx';
import { OMP } from '@turing-weique/office-meta-parser';

const json = toDocxJSON(doc);              // → DocxDocument object
const str  = docx.toJSONString(doc, 2);    // → JSON string
await docx.saveJSON(doc, 'output.json');   // → Write file

// OMP way
OMP.docx.toJSON(doc);
OMP.docx.saveJSON(doc, 'output.json');
docx.toJSON(doc);
docx.toJSONString(doc);
docx.saveJSON(doc, 'output.json');
```

#### Comment Operations

```typescript
import { addComment, removeComment, listComments, getCommentText, markCommentDone, markCommentUndone } from 'office-meta-parser/docx';

// Add a comment anchored to a TextRun
const comment = addComment(doc, run, 'Alice', 'This needs revision');
// run.commentId is set automatically

// List all comments
const comments = listComments(doc);
// [{ comment, blockIndex, runIndices }, ...]

// Get comment plain text
const text = getCommentText(doc, comment.id);  // → 'This needs revision'

// Mark done / undone
markCommentDone(doc, comment.id);
markCommentUndone(doc, comment.id);

// Remove comment (also clears commentId on the run)
removeComment(doc, comment.id);
```

#### Revision (Track Changes) Operations

```typescript
import {
  markInsert, markDelete, addFormatChange, clearRevision,
  listRevisions, hasPendingRevisions,
  acceptAllInserts, acceptAllDeletes, rejectAllInserts, rejectAllDeletes,
} from 'office-meta-parser/docx';

// Mark insert / delete revisions
markInsert(run, 'Alice', '2024-01-01T00:00:00Z');
markDelete(run, 'Alice');

// Add format change record
addFormatChange(doc, 'Bob');

// Query revisions
hasPendingRevisions(doc);          // → boolean
const revisions = listRevisions(doc);  // [{ revision, blockIndex, runIndex, run }, ...]

// Accept / Reject
acceptAllInserts(doc);   // → count (removes markup, keeps text)
acceptAllDeletes(doc);   // → count (removes marked runs)
rejectAllInserts(doc);   // → count (removes marked runs)
rejectAllDeletes(doc);   // → count (removes markup, keeps text)

// Clear revision markup on a single run
clearRevision(run);
```

---

### XLSX (`@turing-weique/office-meta-parser/xlsx`)

#### Parse / Serialize

| Function | Signature | Description |
|----------|-----------|-------------|
| `parseXlsx` | `(buffer: ArrayBuffer) → Promise<{raw, semantic}>` | Decode |
| `serializeXlsx` | `(wb: XlsxWorkbook) → Promise<ArrayBuffer>` | Encode (with validation) |
| `loadXlsx` | `(path: string) → Promise<{raw, semantic}>` | Load from file |
| `saveXlsx` | `(wb: XlsxWorkbook, path: string) → Promise<void>` | Save to file |
| `writeXlsxToStream` | `(wb, stream: Writable) → Promise<void>` | Write to stream |
| `createXlsx` | `(options?) → XlsxWorkbook` | Create empty workbook |
| `validateXlsx` | `(wb: XlsxWorkbook) → ValidationIssue[]` | Validate |

#### Metadata Update

```typescript
import { xlsx } from 'office-meta-parser/xlsx';
import { OMP } from '@turing-weique/office-meta-parser';

// OMP way
OMP.xlsx.updateTitle(wb, 'Sales Report');
OMP.xlsx.updateCreator(wb, 'Bob');

// Namespace way
xlsx.updateTitle(wb, 'Sales Report');
```

#### JSON Export

```typescript
import { xlsx } from 'office-meta-parser/xlsx';
import { OMP } from '@turing-weique/office-meta-parser';

xlsx.toJSON(wb);
xlsx.saveJSON(wb, 'workbook.json');

// OMP way
OMP.xlsx.toJSON(wb);
```

#### Comment Operations

```typescript
import { addComment, removeComment, listComments, listSheetComments, getCommentText, updateComment } from 'office-meta-parser/xlsx';

// Add a comment by cell reference
addComment(wb, 0, 'A1', 'Alice', 'Needs revision');
addComment(wb, 0, 'B2', 'Bob', 'Incorrect data', [{ text: 'Rich text', bold: true }]);

// List comments
const all = listComments(wb);               // Across all sheets
const sheet0 = listSheetComments(wb, 0);    // Specific sheet

// Get / Update
getCommentText(wb, 0, 'A1');                // → 'Needs revision'
updateComment(wb, 0, 'A1', 'Revised');

// Remove
removeComment(wb, 0, 'B2');
```

---

### PPTX (`@turing-weique/office-meta-parser/pptx`)

#### Parse / Serialize

| Function | Signature | Description |
|----------|-----------|-------------|
| `parsePptx` | `(buffer: ArrayBuffer) → Promise<{raw, semantic}>` | Decode |
| `serializePptx` | `(pres: PptxPresentation) → Promise<ArrayBuffer>` | Encode (with validation) |
| `loadPptx` | `(path: string) → Promise<{raw, semantic}>` | Load from file |
| `savePptx` | `(pres: PptxPresentation, path: string) → Promise<void>` | Save to file |
| `writePptxToStream` | `(pres, stream: Writable) → Promise<void>` | Write to stream |
| `createPptx` | `(options?) → PptxPresentation` | Create empty presentation |
| `validatePptx` | `(pres: PptxPresentation) → ValidationIssue[]` | Validate |

#### Metadata Update

```typescript
import { pptx } from 'office-meta-parser/pptx';
import { OMP } from '@turing-weique/office-meta-parser';

// OMP way
OMP.pptx.updateTitle(pres, 'Product Intro');
OMP.pptx.updateCreator(pres, 'Charlie');

// Namespace way
pptx.updateTitle(pres, 'Product Intro');
```

#### JSON Export

```typescript
import { pptx } from 'office-meta-parser/pptx';
import { OMP } from '@turing-weique/office-meta-parser';

pptx.toJSON(pres);
pptx.saveJSON(pres, 'presentation.json');

// OMP way
OMP.pptx.toJSON(pres);
```

#### Comment Operations

```typescript
import { addComment, removeComment, listComments, listSlideComments, getCommentText } from 'office-meta-parser/pptx';

// Add a comment with optional position coordinates
addComment(pres, 0, 'Reviewer', 'Title needs revision', 100, 200);

// List comments
const all = listComments(pres);               // Across all slides
const slide0 = listSlideComments(pres, 0);    // Specific slide

// Get text
getCommentText(pres, 0, '1');                 // → 'Title needs revision'

// Remove
removeComment(pres, 0, '1');
```

---

### Core Module (`@turing-weique/office-meta-parser/core`)

#### File I/O

| Function | Signature | Description |
|----------|-----------|-------------|
| `loadFromFile` | `(path: string) → Promise<ArrayBuffer>` | Read file |
| `saveToFile` | `(buffer: ArrayBuffer, path: string) → Promise<void>` | Write file |
| `writeToStream` | `(buffer, stream: Writable) → Promise<void>` | Write to stream |
| `toBuffer` | `(buffer: ArrayBuffer) → Buffer` | Convert to Node.js Buffer |
| `toJSON` | `<T>(semantic: T) → T` | Semantic model → JSON object |
| `toJSONString` | `<T>(semantic: T, space?) → string` | Semantic model → JSON string |
| `saveToJSON` | `<T>(semantic: T, path: string, space?) → Promise<void>` | Semantic model → JSON file |

#### Format Detection & Validation

```typescript
import { detectFormat, validate } from '@turing-weique/office-meta-parser';

// Detect format
const format = await detectFormat(buffer); // 'docx' | 'xlsx' | 'pptx' | null

// Unified validation (auto-detect + parse + validate)
const result = await validate(buffer);
// { format: 'docx', issues: [...], valid: true }
```

#### Metadata Update (Generic)

```typescript
import { updateTitle, updateCreator, createMetaOps } from '@turing-weique/office-meta-parser';

// Use generic functions directly — works on any object with meta
updateTitle(doc, 'New Title');
updateCreator(wb, 'Alice');

// createMetaOps factory — generate full operations for custom types
import type { DocumentMeta } from '@turing-weique/office-meta-parser';

interface MyDoc { meta: DocumentMeta; /* ... */ }
const myOps = createMetaOps<MyDoc>();

myOps.updateTitle(myDoc, 'New Title');
myOps.updateCreator(myDoc, 'Alice');
myOps.toJSON(myDoc);
myOps.saveJSON(myDoc, 'output.json');
// 10 operations total: 7 metadata + toJSON / toJSONString / saveJSON
```

#### XML / ZIP

| Function | Description |
|----------|-------------|
| `parseXml(xml)` | XML string → ParsedNode tree |
| `serializeXml(node)` | ParsedNode tree → XML string |
| `parseRels(xml)` | Parse .rels relationship file |
| `serializeRels(rels)` | Serialize .rels |
| `parseContentTypes(xml)` | Parse [Content_Types].xml |
| `serializeContentTypes(cts)` | Serialize Content_Types |
| `unzip(buffer)` | Decompress ZIP → ZipEntry[] |
| `zip(entries)` | Compress to ZIP |

#### Metadata Parse / Serialize

| Function | Description |
|----------|-------------|
| `parseMeta(node)` | Parse core.xml metadata |
| `serializeMeta(meta)` | Serialize core.xml |
| `parseAppMeta(xml)` | Parse app.xml |
| `serializeAppMeta(meta)` | Serialize app.xml |
| `parseCustomProperties(xml)` | Parse custom.xml |
| `serializeCustomProperties(props)` | Serialize custom.xml |

---

### Error Handling

```typescript
import { ValidationError, FormatError } from '@turing-weique/office-meta-parser';

try {
  await serializeDocx(doc);
} catch (e) {
  if (e instanceof ValidationError) {
    // Validation failed: e.issues contains the error list
    for (const issue of e.issues) {
      console.error(`[${issue.level}] ${issue.path}: ${issue.message}`);
    }
  }
}

// FormatError for unsupported formats
const err = new FormatError('Unsupported format', 'pdf');
```

---

### Validation Rules

Each format's validator includes **required** (error — blocks serialization) and **optional** (warning — prints to stderr) rules.

#### DOCX

| Level | Rule |
|-------|------|
| error | body.blocks must not be empty |
| error | commentId must reference an existing comment |
| error | bookmarkStart/bookmarkEnd must be paired |
| error | hyperlink/image must have a relationshipId |
| error | numbering.numId reference must exist |
| error | header/footer id must not be duplicated |
| error | table.rows must not be empty |
| warning | comment content must not be empty |
| warning | comments defined but no reference in body |

#### XLSX

| Level | Rule |
|-------|------|
| error | sheets must not be empty |
| error | sheet.name must not be empty |
| error | sharedString reference must be in bounds |
| error | mergedCells range must be valid |
| error | hyperlinks must have ref and url |
| error | table must have ref and displayName |

#### PPTX

| Level | Rule |
|-------|------|
| error | theme must not be missing |
| error | masters/layouts must not be empty |
| error | slides must not be empty |
| error | image must have a relationshipId |
| error | table.rows must not be empty |
| warning | layout reference must exist |
| warning | slideSize missing |

---

## Data Models

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
  fontSize?: number;       // Half-point value (e.g. 56 = 28pt)
  color?: string;          // RGB hex
  fontFamily?: string;
  highlight?: string;
  commentId?: string;      // Associated comment
}

interface Comment {
  id: string;
  author: string;
  date: string;
  content: Paragraph[];    // Comment body (structured paragraphs)
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
  columnWidths: number[];
  rowHeights: number[];
  hyperlinks: Hyperlink[];
  autoFilter?: AutoFilter;
  dataValidations?: DataValidation[];
  conditionalFormats?: ConditionalFormat[];
  frozenPanes?: FrozenPanes;
  comments?: SheetComment[];
  tables?: ExcelTable[];
}

interface Cell {
  type: 'string' | 'sharedString' | 'number' | 'boolean' | 'formula' | 'error';
  value: string | number | boolean | null;
  formula?: string;
}

interface SheetComment {
  ref: string;           // Cell reference, e.g. 'A1'
  authorId: number;
  text: string;
  richText?: RichTextRun[];
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
  layout?: string;
  transition?: Transition;
  notes?: string;
  animations?: Animation[];
  comments?: SlideComment[];
}

type SlideElement = TextShape | ImageShape | GroupShape | TableShape | MediaShape;

interface SlideComment {
  id: string;
  authorId: number;
  authorName: string;
  text: string;
  date?: string;
  position?: { x: number; y: number };
  replies?: SlideComment[];
}
```

---

## Testing

```bash
npm test              # Run all tests (89 files / 507 cases)
npm run test:codec    # Codec integration tests
npm run typecheck     # Type check
```

## Dependencies

| Package | Purpose |
|---------|---------|
| `jszip` | ZIP compression/decompression |
| `fast-xml-parser` | XML parsing and serialization |

## License

[Apache License 2.0](./LICENSE)
