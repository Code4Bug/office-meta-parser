#!/usr/bin/env tsx
/**
 * 测试 fixture 生成脚本
 * 基于项目的元数据编码能力，构建覆盖所有情况的 docx / pptx / xlsx 测试文件
 *
 * 用法: tsx tests/scripts/generate-fixtures.ts
 */

import { mkdirSync, statSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

import { saveDocx, validateDocx } from '../../src/docx/index.js';
import { saveXlsx, validateXlsx } from '../../src/xlsx/index.js';
import { savePptx, validatePptx } from '../../src/pptx/index.js';
import { formatValidationReport } from '../../src/core/validate.js';

import type { DocxDocument, DocxBlock, Paragraph, Table, TextRun, Comment, Revision, Header, Footer, NumberingDefinitions, Footnote, FontEntry, DocumentSettings, CommentExtended, Person } from '../../src/docx/types.js';
import type { XlsxWorkbook, Sheet, Cell, MergedCell, SharedStringEntry } from '../../src/xlsx/types.js';
import type { PptxPresentation, Slide, SlideElement, TextShape, SlideMaster, SlideLayout, Placeholder } from '../../src/pptx/types.js';
import type { DocumentMeta } from '../../src/core/meta.js';
import type { AppMeta } from '../../src/core/app-meta.js';
import type { CustomProperty } from '../../src/core/custom-meta.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const OUTPUT_DIR = join(__dirname, '..', 'fixtures', 'generated');
mkdirSync(OUTPUT_DIR, { recursive: true });

// ─── 通用元数据 ───────────────────────────────────────────────────────────────

const FULL_CORE_META: DocumentMeta = {
  title: '测试文档标题',
  subject: '测试主题',
  creator: '张三',
  description: '这是一个用于测试的文档描述',
  keywords: '测试,元数据,office',
  lastModifiedBy: '李四',
  created: '2024-01-15T08:30:00Z',
  modified: '2024-06-20T14:45:00Z',
  revision: '7',
  category: '测试分类',
};

const FULL_APP_META: AppMeta = {
  template: 'Normal.dotm',
  totalTime: 120,
  pages: 5,
  words: 3500,
  characters: 18000,
  charactersWithSpaces: 21000,
  application: 'Microsoft Office Word',
  docSecurity: 0,
  scaleCrop: false,
  company: '测试公司',
  linksUpToDate: false,
  sharedDoc: false,
  hyperlinksChanged: false,
  appVersion: '16.0000',
};

const FULL_CUSTOM_PROPS: CustomProperty[] = [
  { name: 'CustomString', value: '自定义字符串值', type: 'lpwstr' },
  { name: 'CustomInt', value: 42, type: 'i4' },
  { name: 'CustomFloat', value: 3.14159, type: 'r8' },
  { name: 'CustomBoolTrue', value: true, type: 'bool' },
  { name: 'CustomBoolFalse', value: false, type: 'bool' },
  { name: 'CustomDate', value: '2024-03-15T10:00:00Z', type: 'date' },
];

// ─── 工具函数 ─────────────────────────────────────────────────────────────────

async function saveFixture(name: string, doc: DocxDocument | XlsxWorkbook | PptxPresentation, format: 'docx' | 'xlsx' | 'pptx') {
  const path = join(OUTPUT_DIR, `${name}.${format}`);
  if (format === 'docx') await saveDocx(doc as DocxDocument, path);
  else if (format === 'xlsx') await saveXlsx(doc as XlsxWorkbook, path);
  else await savePptx(doc as PptxPresentation, path);

  const size = statSync(path).size;
  console.log(`  ✓ ${name}.${format} (${size} bytes)`);
}

function p(text: string, props?: Partial<TextRun>): Paragraph {
  return {
    type: 'paragraph',
    runs: [{ text, ...props }],
  };
}

function styledParagraph(runs: TextRun[], style?: string): Paragraph {
  return { type: 'paragraph', runs, style };
}

function emptyDocx(): DocxDocument {
  return {
    meta: {},
    styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
    body: { blocks: [] },
  };
}

function emptyXlsx(): XlsxWorkbook {
  return {
    meta: {},
    sheets: [],
    styles: { cellStyles: [], fonts: [], fills: [], borders: [], numberFormats: [] },
    sharedStrings: [],
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCX Fixtures
// ═══════════════════════════════════════════════════════════════════════════════

async function docxMetaFull() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: FULL_CORE_META,
    appMeta: FULL_APP_META,
    customProperties: FULL_CUSTOM_PROPS,
    body: { blocks: [p('包含完整元数据的测试文档')] },
  };
  await saveFixture('docx-meta-full', doc, 'docx');
}

async function docxMetaCoreOnly() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: {
      title: '仅核心属性',
      subject: '',
      creator: '',
      description: '',
      keywords: '',
      lastModifiedBy: '',
      created: '',
      modified: '',
      revision: '',
      category: '',
    },
    body: { blocks: [p('仅包含核心元数据')] },
  };
  await saveFixture('docx-meta-core-only', doc, 'docx');
}

async function docxMetaAppOnly() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: {},
    appMeta: {
      template: 'Normal.dotm',
      totalTime: 45,
      pages: 1,
      words: 100,
      characters: 500,
      charactersWithSpaces: 600,
      application: 'Microsoft Office Word',
      docSecurity: 0,
      scaleCrop: false,
      company: '仅应用属性公司',
      linksUpToDate: false,
      sharedDoc: false,
      hyperlinksChanged: false,
      appVersion: '16.0000',
    },
    body: { blocks: [p('仅包含应用元数据')] },
  };
  await saveFixture('docx-meta-app-only', doc, 'docx');
}

async function docxMetaCustomOnly() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: {},
    customProperties: [
      { name: 'StringProp', value: 'Hello World', type: 'lpwstr' },
      { name: 'IntProp', value: 100, type: 'i4' },
      { name: 'FloatProp', value: 2.71828, type: 'r8' },
      { name: 'BoolProp', value: true, type: 'bool' },
      { name: 'DateProp', value: '2024-12-25T00:00:00Z', type: 'date' },
    ],
    body: { blocks: [p('仅包含自定义属性')] },
  };
  await saveFixture('docx-meta-custom-only', doc, 'docx');
}

async function docxContentParagraphs() {
  const runs: TextRun[] = [
    { text: '普通文本 ' },
    { text: '粗体', bold: true },
    { text: ' ' },
    { text: '斜体', italic: true },
    { text: ' ' },
    { text: '下划线', underline: true },
    { text: ' ' },
    { text: '删除线', strike: true },
    { text: ' ' },
    { text: '上标', superscript: true },
    { text: ' ' },
    { text: '下标', subscript: true },
    { text: ' ' },
    { text: '大写', caps: true },
    { text: ' ' },
    { text: '小型大写', smallCaps: true },
    { text: ' ' },
    { text: '双删除线', dstrike: true },
    { text: ' ' },
    { text: '隐藏文字', vanish: true },
  ];

  const fontSizeRuns: TextRun[] = [
    { text: '小字8pt ', fontSize: 8 },
    { text: '正常12pt ', fontSize: 12 },
    { text: '大字24pt ', fontSize: 24 },
    { text: '特大36pt', fontSize: 36 },
  ];

  const colorRuns: TextRun[] = [
    { text: '红色文字 ', color: 'FF0000' },
    { text: '蓝色文字 ', color: '0000FF' },
    { text: '绿色高亮 ', highlight: 'green' },
    { text: '黄色底纹 ', shadingColor: 'FFFF00' },
  ];

  const spacingRuns: TextRun[] = [
    { text: '加宽间距 ', characterSpacing: 100 },
    { text: '紧缩间距 ', characterSpacing: -50 },
    { text: '字距调整', kern: 10 },
  ];

  const advancedRuns: TextRun[] = [
    { text: '着重号', emphasis: 'dot' },
    { text: ' ' },
    { text: '上移3pt', verticalPosition: 3 },
    { text: ' ' },
    { text: '字体指定', fontFamily: 'SimSun' },
  ];

  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '段落与文字格式测试' },
    body: {
      blocks: [
        p('这是一个没有格式的普通段落。'),
        { type: 'paragraph', runs },
        styledParagraph(fontSizeRuns),
        styledParagraph(colorRuns),
        styledParagraph(spacingRuns),
        styledParagraph(advancedRuns),
        {
          type: 'paragraph',
          runs: [{ text: '居中对齐段落' }],
          properties: { alignment: 'center' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '右对齐段落' }],
          properties: { alignment: 'right' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '两端对齐段落，这是一段较长的文本用于演示两端对齐效果。' }],
          properties: { alignment: 'justify' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '带缩进的段落' }],
          properties: {
            indent: { left: 720, right: 360, firstLine: 480 },
          },
        },
        {
          type: 'paragraph',
          runs: [{ text: '段前段后间距' }],
          properties: {
            spacing: { before: 240, after: 120, line: 360, lineRule: 'auto' },
          },
        },
        {
          type: 'paragraph',
          runs: [{ text: '保持与下段同页 / 段前分页' }],
          properties: { keepNext: true, pageBreakBefore: true },
        },
        {
          type: 'paragraph',
          runs: [{ text: '大纲级别2' }],
          properties: { outlineLevel: 2 },
        },
        {
          type: 'paragraph',
          runs: [{ text: '带边框和底纹的段落' }],
          properties: {
            border: {
              top: { style: 'single', size: 4, color: '0000FF' },
              bottom: { style: 'single', size: 4, color: '0000FF' },
            },
            shading: { fill: 'E0E0E0', color: 'auto', pattern: 'clear' },
          },
        },
        {
          type: 'paragraph',
          runs: [{ text: '带制表位的段落' }],
          properties: {
            tabs: [
              { position: 1440, alignment: 'left' },
              { position: 4320, alignment: 'center', leader: 'dot' },
              { position: 7200, alignment: 'right', leader: 'underscore' },
            ],
          },
        },
        {
          type: 'paragraph',
          runs: [{ text: '修订插入文本', revisionType: 'insert', revisionAuthor: '王五', revisionDate: '2024-05-01T10:00:00Z' }],
        },
        {
          type: 'paragraph',
          runs: [{ text: '修订删除文本', revisionType: 'delete', revisionAuthor: '赵六', revisionDate: '2024-05-02T11:00:00Z' }],
        },
      ],
    },
  };
  await saveFixture('docx-content-paragraphs', doc, 'docx');
}

async function docxContentTables() {
  const table1: Table = {
    type: 'table',
    rows: [
      {
        cells: [
          { blocks: [p('表头A')] },
          { blocks: [p('表头B')] },
          { blocks: [p('表头C')] },
        ],
      },
      {
        cells: [
          { blocks: [p('数据1')] },
          { blocks: [p('数据2')] },
          { blocks: [p('数据3')] },
        ],
      },
    ],
    properties: {
      style: 'TableGrid',
      width: 5000,
      layout: 'autofit',
      borders: {
        top: { style: 'single', size: 4, color: '000000' },
        bottom: { style: 'single', size: 4, color: '000000' },
        left: { style: 'single', size: 4, color: '000000' },
        right: { style: 'single', size: 4, color: '000000' },
        insideHorizontal: { style: 'single', size: 4, color: '000000' },
        insideVertical: { style: 'single', size: 4, color: '000000' },
      },
      cellMarginTop: 50,
      cellMarginLeft: 100,
      cellMarginBottom: 50,
      cellMarginRight: 100,
    },
  };

  const table2: Table = {
    type: 'table',
    rows: [
      {
        cells: [
          { blocks: [p('合并单元格')] },
          {
            blocks: [p('垂直合并')],
            properties: { verticalMerge: 'restart' },
          },
        ],
      },
      {
        cells: [
          { blocks: [p('行2列1')] },
          {
            blocks: [p('')],
            properties: { verticalMerge: 'continue' },
          },
        ],
      },
    ],
    properties: {
      width: 4000,
      layout: 'fixed',
      gridColumns: [2000, 2000],
    },
  };

  const table3: Table = {
    type: 'table',
    rows: [
      {
        cells: [
          {
            blocks: [p('跨列')],
            properties: { gridSpan: 2, verticalAlign: 'center' },
          },
        ],
      },
      {
        cells: [
          { blocks: [p('A')], properties: { width: 1500, noWrap: true } },
          { blocks: [p('B')], properties: { width: 1500 } },
        ],
      },
    ],
  };

  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '表格测试' },
    body: {
      blocks: [
        p('基本表格：'),
        table1,
        p('合并单元格表格：'),
        table2,
        p('跨列与垂直对齐：'),
        table3,
      ],
    },
  };
  await saveFixture('docx-content-tables', doc, 'docx');
}

async function docxContentHyperlinks() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '超链接测试' },
    body: {
      blocks: [
        p('以下包含超链接：'),
        {
          type: 'hyperlink',
          relationshipId: 'rId10',
          url: 'https://example.com',
          runs: [{ text: '示例网站' }],
          tooltip: '点击访问示例网站',
        },
        {
          type: 'hyperlink',
          relationshipId: 'rId11',
          url: 'mailto:test@example.com',
          runs: [{ text: '发送邮件', bold: true }],
        },
        {
          type: 'hyperlink',
          relationshipId: 'rId12',
          url: 'https://example.org/path?query=1',
          runs: [
            { text: '带' },
            { text: '格式', italic: true, color: '0000FF' },
            { text: '的超链接' },
          ],
        },
      ],
    },
  };
  await saveFixture('docx-content-hyperlinks', doc, 'docx');
}

async function docxContentBookmarks() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '书签测试' },
    body: {
      blocks: [
        { type: 'bookmarkStart', id: '1', name: 'bookmark_start' },
        p('这是一个书签标记的段落。'),
        { type: 'bookmarkEnd', id: '1' },
        p('普通段落。'),
        { type: 'bookmarkStart', id: '2', name: 'section_ref' },
        { type: 'bookmarkStart', id: '3', name: 'nested_bookmark' },
        p('嵌套书签区域。'),
        { type: 'bookmarkEnd', id: '3' },
        { type: 'bookmarkEnd', id: '2' },
      ],
    },
  };
  await saveFixture('docx-content-bookmarks', doc, 'docx');
}

async function docxContentCommentsRevisions() {
  const comments: Comment[] = [
    {
      id: '1',
      author: '审阅者A',
      date: '2024-03-10T09:00:00Z',
      content: [p('这是第一条评论。')],
      initials: 'A',
    },
    {
      id: '2',
      author: '审阅者B',
      date: '2024-03-11T14:30:00Z',
      content: [
        p('这是第二条评论，'),
        p('包含多个段落。'),
      ],
      initials: 'B',
    },
    {
      id: '3',
      author: '审阅者C',
      date: '2024-03-12T16:00:00Z',
      content: [p('第三条批注，带粗体文字。')],
    },
  ];

  const trackChanges: Revision[] = [
    { type: 'insert', author: '编辑A', date: '2024-04-01T10:00:00Z' },
    { type: 'delete', author: '编辑B', date: '2024-04-02T11:00:00Z' },
    { type: 'formatChange', author: '编辑C', date: '2024-04-03T12:00:00Z' },
  ];

  const commentExts: CommentExtended[] = [
    { paraId: '00000001', done: false },
    { paraId: '00000002', done: true },
    { paraId: '00000003' },
  ];

  const people: Person[] = [
    { author: '审阅者A', userId: 'user-a-id', providerId: 'provider-1' },
    { author: '审阅者B', userId: 'user-b-id' },
    { author: '审阅者C' },
  ];

  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '批注与修订测试' },
    comments,
    trackChanges,
    commentExts,
    people,
    body: {
      blocks: [
        { type: 'paragraph', runs: [{ text: '普通段落。', commentId: '1' }] },
        {
          type: 'paragraph',
          runs: [
            { text: '带', revisionType: 'insert', revisionAuthor: '编辑A', revisionDate: '2024-04-01T10:00:00Z' },
            { text: '修订', revisionType: 'insert', revisionAuthor: '编辑A', revisionDate: '2024-04-01T10:00:00Z', commentId: '2' },
            { text: '标记', revisionType: 'insert', revisionAuthor: '编辑A', revisionDate: '2024-04-01T10:00:00Z' },
            { text: '的段落。' },
          ],
        },
        {
          type: 'paragraph',
          runs: [
            { text: '被删除的文本', revisionType: 'delete', revisionAuthor: '编辑B', revisionDate: '2024-04-02T11:00:00Z' },
            { text: '保留的文本', commentId: '3' },
          ],
        },
      ],
    },
  };
  await saveFixture('docx-content-comments-revisions', doc, 'docx');
}

async function docxContentHeadersFooters() {
  const headers: Header[] = [
    {
      id: 'rId20',
      type: 'default',
      content: [p('这是默认页眉')],
    },
    {
      id: 'rId21',
      type: 'first',
      content: [p('首页页眉')],
    },
  ];

  const footers: Footer[] = [
    {
      id: 'rId22',
      type: 'default',
      content: [
        {
          type: 'paragraph',
          runs: [
            { text: '第 ' },
            { text: '页码', field: { instruction: 'PAGE' } },
            { text: ' 页' },
          ],
          properties: { alignment: 'center' },
        },
      ],
    },
    {
      id: 'rId23',
      type: 'even',
      content: [p('偶数页页脚')],
    },
  ];

  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '页眉页脚测试' },
    headers,
    footers,
    body: {
      blocks: [p('包含页眉页脚的文档')],
      sectionProperties: {
        pageWidth: 11906,
        pageHeight: 16838,
        marginTop: 1440,
        marginRight: 1800,
        marginBottom: 1440,
        marginLeft: 1800,
        headerMargin: 720,
        footerMargin: 720,
        gutter: 0,
        orientation: 'portrait',
        titlePage: true,
        evenAndOddHeaders: true,
        pageNumberFormat: 'decimal',
        pageNumberStart: 1,
        verticalAlign: 'top',
        headerReferenceId: 'rId20',
        headerReferenceType: 'default',
        footerReferenceId: 'rId22',
        footerReferenceType: 'default',
      },
    },
  };
  await saveFixture('docx-content-headers-footers', doc, 'docx');
}

async function docxContentNumbering() {
  const numbering: NumberingDefinitions = {
    abstractNums: [
      {
        id: '0',
        levels: [
          { level: 0, numFmt: 'decimal', lvlText: '%1.', start: 1, indent: { left: 720, hanging: 360 }, alignment: 'left' },
          { level: 1, numFmt: 'lowerLetter', lvlText: '%2)', start: 1, indent: { left: 1440, hanging: 360 } },
          { level: 2, numFmt: 'lowerRoman', lvlText: '%3.', start: 1, indent: { left: 2160, hanging: 360 } },
        ],
      },
      {
        id: '1',
        levels: [
          { level: 0, numFmt: 'bullet', lvlText: '•', start: 1, indent: { left: 720, hanging: 360 } },
          { level: 1, numFmt: 'bullet', lvlText: '◦', start: 1, indent: { left: 1440, hanging: 360 } },
        ],
      },
    ],
    nums: [
      { id: '1', abstractNumId: '0' },
      { id: '2', abstractNumId: '1' },
    ],
  };

  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '编号列表测试' },
    numbering,
    body: {
      blocks: [
        p('有序列表：'),
        {
          type: 'paragraph',
          runs: [{ text: '第一项' }],
          numbering: { level: 0, numId: '1', format: 'decimal', text: '1.' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '子项a' }],
          numbering: { level: 1, numId: '1', format: 'lowerLetter', text: 'a)' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '子项b' }],
          numbering: { level: 1, numId: '1', format: 'lowerLetter', text: 'b)' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '第二项' }],
          numbering: { level: 0, numId: '1', format: 'decimal', text: '2.' },
        },
        p('无序列表：'),
        {
          type: 'paragraph',
          runs: [{ text: '项目A' }],
          numbering: { level: 0, numId: '2', format: 'bullet', text: '•' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '子项目' }],
          numbering: { level: 1, numId: '2', format: 'bullet', text: '◦' },
        },
        {
          type: 'paragraph',
          runs: [{ text: '项目B' }],
          numbering: { level: 0, numId: '2', format: 'bullet', text: '•' },
        },
      ],
    },
  };
  await saveFixture('docx-content-numbering', doc, 'docx');
}

async function docxContentFootnotesEndnotes() {
  const footnotes: Footnote[] = [
    { id: '1', content: [p('这是第一个脚注。')] },
    { id: '2', content: [p('这是第二个脚注，包含'), p('多个段落。')] },
  ];

  const endnotes: Footnote[] = [
    { id: '1', content: [p('这是第一个尾注。')] },
    { id: '2', content: [p('这是第二个尾注。')] },
  ];

  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '脚注尾注测试' },
    footnotes,
    endnotes,
    body: {
      blocks: [
        p('正文包含脚注和尾注引用。'),
        p('第二个段落也有引用。'),
      ],
    },
  };
  await saveFixture('docx-content-footnotes-endnotes', doc, 'docx');
}

async function docxContentStyles() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '样式测试' },
    styles: {
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          isDefault: false,
          uiPriority: 9,
          semiHidden: false,
          unhideWhenUsed: false,
          qFormat: true,
          properties: {
            spacing: { before: 240, after: 120 },
            outlineLevel: 0,
          },
          runProperties: {
            bold: true,
            fontSize: 28,
            color: '2F5496',
            fontFamily: '微软雅黑',
          },
        },
        {
          id: 'Heading2',
          name: 'heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          uiPriority: 9,
          qFormat: true,
          properties: {
            spacing: { before: 200, after: 80 },
            outlineLevel: 1,
          },
          runProperties: {
            bold: true,
            fontSize: 24,
            color: '2F5496',
          },
        },
        {
          id: 'Normal',
          name: 'Normal',
          isDefault: true,
          properties: {},
          runProperties: { fontSize: 12, fontFamily: 'Calibri' },
        },
      ],
      characterStyles: [
        {
          id: 'Strong',
          name: 'Strong',
          basedOn: 'DefaultParagraphFont',
          properties: { bold: true },
        },
        {
          id: 'Emphasis',
          name: 'Emphasis',
          basedOn: 'DefaultParagraphFont',
          properties: { italic: true },
        },
      ],
      tableStyles: [
        { id: 'TableGrid', name: 'Table Grid', basedOn: 'NormalTable' },
      ],
      latentStyles: [
        { name: 'Heading1', uiPriority: 9, semiHidden: false, unhideWhenUsed: false, qFormat: true },
        { name: 'Heading2', uiPriority: 9, qFormat: true },
      ],
    },
    body: {
      blocks: [
        { type: 'paragraph', runs: [{ text: '标题一' }], style: 'Heading1' },
        { type: 'paragraph', runs: [{ text: '正文内容。' }] },
        { type: 'paragraph', runs: [{ text: '标题二' }], style: 'Heading2' },
        { type: 'paragraph', runs: [{ text: '更多正文。' }] },
        {
          type: 'paragraph',
          runs: [
            { text: '普通' },
            { text: '加粗', properties: { bold: true } },
            { text: '和' },
            { text: '强调', properties: { italic: true } },
            { text: '文字。' },
          ],
        },
      ],
    },
  };
  await saveFixture('docx-content-styles', doc, 'docx');
}

async function docxContentSettingsFontsTheme() {
  const fonts: FontEntry[] = [
    { name: 'Calibri', altName: 'Helvetica', charset: '00', family: 'swiss', pitch: 'variable' },
    { name: 'SimSun', altName: '宋体', charset: '86', family: 'auto', pitch: 'variable' },
    { name: 'SimHei', altName: '黑体', charset: '86', family: 'modern', pitch: 'variable' },
    { name: 'Times New Roman', charset: '00', family: 'roman', pitch: 'variable' },
  ];

  const settings: DocumentSettings = {
    defaultTabStop: 720,
    zoom: 100,
    compatibilityMode: 15,
    evenAndOddHeaders: true,
    documentProtection: false,
    characterSpacingControl: 'doNotCompress',
  };

  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '设置字体主题测试' },
    fonts,
    settings,
    body: { blocks: [p('包含字体表和设置的文档')] },
  };
  await saveFixture('docx-settings-fonts-theme', doc, 'docx');
}

async function docxContentFields() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: { title: '域代码测试' },
    body: {
      blocks: [
        {
          type: 'paragraph',
          runs: [
            { text: '页码: ' },
            { text: '1', field: { instruction: 'PAGE', result: '1' } },
          ],
        },
        {
          type: 'paragraph',
          runs: [
            { text: '总页数: ' },
            { text: '10', field: { instruction: 'NUMPAGES', result: '10' } },
          ],
        },
        {
          type: 'paragraph',
          runs: [
            { text: '日期: ' },
            { text: '2024-06-20', field: { instruction: 'DATE \\@ "yyyy-MM-dd"', result: '2024-06-20' } },
          ],
        },
        {
          type: 'paragraph',
          runs: [
            { text: '文件名: ' },
            { text: 'test.docx', field: { instruction: 'FILENAME', result: 'test.docx' } },
          ],
        },
      ],
    },
  };
  await saveFixture('docx-content-fields', doc, 'docx');
}

async function docxAllInOne() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: FULL_CORE_META,
    appMeta: FULL_APP_META,
    customProperties: FULL_CUSTOM_PROPS,
    styles: {
      paragraphStyles: [
        { id: 'Normal', name: 'Normal', isDefault: true, runProperties: { fontSize: 12 } },
        { id: 'Heading1', name: 'heading 1', basedOn: 'Normal', qFormat: true, runProperties: { bold: true, fontSize: 28 } },
      ],
      characterStyles: [
        { id: 'Strong', name: 'Strong', properties: { bold: true } },
      ],
      tableStyles: [],
    },
    fonts: [
      { name: 'Calibri', charset: '00', family: 'swiss', pitch: 'variable' },
      { name: 'SimSun', charset: '86', family: 'auto', pitch: 'variable' },
    ],
    settings: { defaultTabStop: 720, zoom: 100, compatibilityMode: 15 },
    numbering: {
      abstractNums: [
        {
          id: '0',
          levels: [{ level: 0, numFmt: 'decimal', lvlText: '%1.', start: 1, indent: { left: 720, hanging: 360 } }],
        },
      ],
      nums: [{ id: '1', abstractNumId: '0' }],
    },
    comments: [
      { id: '1', author: '审阅者', date: '2024-03-10T09:00:00Z', content: [p('综合测试批注')] },
    ],
    commentExts: [{ paraId: '00000001', done: false }],
    people: [{ author: '审阅者', userId: 'reviewer-id', providerId: 'pid' }],
    headers: [{ id: 'rId20', type: 'default', content: [p('综合测试页眉')] }],
    footers: [{ id: 'rId22', type: 'default', content: [p('综合测试页脚')] }],
    body: {
      blocks: [
        { type: 'paragraph', runs: [{ text: '综合测试文档' }], style: 'Heading1' },
        { type: 'paragraph', runs: [{ text: '这是包含所有功能的综合测试文档。', commentId: '1' }] },
        {
          type: 'paragraph',
          runs: [
            { text: '粗体斜体下划线', bold: true, italic: true, underline: true },
            { text: ' ' },
            { text: '红色', color: 'FF0000', fontSize: 16 },
          ],
        },
        {
          type: 'table',
          rows: [
            { cells: [{ blocks: [p('A')] }, { blocks: [p('B')] }] },
            { cells: [{ blocks: [p('1')] }, { blocks: [p('2')] }] },
          ],
          properties: { width: 3000 },
        },
        {
          type: 'hyperlink',
          relationshipId: 'rId10',
          url: 'https://example.com',
          runs: [{ text: '示例链接' }],
        },
        { type: 'bookmarkStart', id: '1', name: '综合书签' },
        p('书签标记的内容。'),
        { type: 'bookmarkEnd', id: '1' },
        {
          type: 'paragraph',
          runs: [{ text: '带编号的项目' }],
          numbering: { level: 0, numId: '1' },
        },
      ],
      sectionProperties: {
        pageWidth: 11906,
        pageHeight: 16838,
        marginTop: 1440,
        marginRight: 1800,
        marginBottom: 1440,
        marginLeft: 1800,
        orientation: 'portrait',
      },
    },
  };
  await saveFixture('docx-all-in-one', doc, 'docx');
}

async function docxEdgeCases() {
  const doc: DocxDocument = {
    ...emptyDocx(),
    meta: {
      title: '',
      subject: '',
      creator: '',
      description: '',
      keywords: '',
      lastModifiedBy: '',
      created: '',
      modified: '',
      revision: '',
      category: '',
    },
    body: {
      blocks: [
        p(''),
        { type: 'paragraph', runs: [] },
        {
          type: 'paragraph',
          runs: [
            { text: '' },
            { text: 'A', fontSize: 1 },
            { text: 'B'.repeat(500) },
          ],
        },
        {
          type: 'paragraph',
          runs: [{ text: '特殊字符: <>&"\'中文日本語한국어' }],
        },
        {
          type: 'paragraph',
          runs: [{ text: '零宽字符: ​‌‍﻿' }],
        },
        {
          type: 'table',
          rows: [
            {
              cells: [
                { blocks: [] },
                { blocks: [p('')] },
              ],
            },
          ],
        },
      ],
    },
  };
  await saveFixture('docx-edge-cases', doc, 'docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// XLSX Fixtures
// ═══════════════════════════════════════════════════════════════════════════════

async function xlsxMetaFull() {
  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: FULL_CORE_META,
    appMeta: FULL_APP_META,
    customProperties: FULL_CUSTOM_PROPS,
    sheets: [{ name: 'Sheet1', cells: [], mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [] }],
  };
  await saveFixture('xlsx-meta-full', wb, 'xlsx');
}

async function xlsxCellTypes() {
  const cells: Cell[][] = [
    // Row 0: headers
    [
      { value: '类型', type: 'sharedString', style: 'header' },
      { value: '值', type: 'sharedString', style: 'header' },
      { value: '说明', type: 'sharedString', style: 'header' },
    ],
    // Row 1: string
    [
      { value: '字符串', type: 'sharedString' },
      { value: 'Hello World', type: 'string' },
      { value: '普通字符串', type: 'sharedString' },
    ],
    // Row 2: number
    [
      { value: '数字', type: 'sharedString' },
      { value: 42, type: 'number' },
      { value: '整数', type: 'sharedString' },
    ],
    // Row 3: float
    [
      { value: '浮点数', type: 'sharedString' },
      { value: 3.14159, type: 'number' },
      { value: '小数', type: 'sharedString' },
    ],
    // Row 4: boolean
    [
      { value: '布尔', type: 'sharedString' },
      { value: true, type: 'boolean' },
      { value: '布尔值', type: 'sharedString' },
    ],
    // Row 5: formula
    [
      { value: '公式', type: 'sharedString' },
      { value: '', type: 'formula', formula: 'B3+B4' },
      { value: '求和公式', type: 'sharedString' },
    ],
    // Row 6: date
    [
      { value: '日期', type: 'sharedString' },
      { value: '2024-06-20', type: 'date' },
      { value: '日期值', type: 'sharedString' },
    ],
    // Row 7: error
    [
      { value: '错误', type: 'sharedString' },
      { value: '#DIV/0!', type: 'error' },
      { value: '除零错误', type: 'sharedString' },
    ],
    // Row 8: rich text
    [
      { value: '富文本', type: 'sharedString' },
      { value: null, type: 'string', richText: [
        { text: '粗体', bold: true },
        { text: '和' },
        { text: '斜体', italic: true, color: 'FF0000', size: 14, font: 'SimSun' },
      ]},
      { value: '带格式文字', type: 'sharedString' },
    ],
  ];

  const sharedStrings: SharedStringEntry[] = [
    { text: '类型' },
    { text: '值' },
    { text: '说明' },
    { text: '字符串' },
    { text: '普通字符串' },
    { text: '数字' },
    { text: '整数' },
    { text: '浮点数' },
    { text: '小数' },
    { text: '布尔' },
    { text: '布尔值' },
    { text: '公式' },
    { text: '求和公式' },
    { text: '日期' },
    { text: '日期值' },
    { text: '错误' },
    { text: '除零错误' },
    { text: '富文本' },
    { text: '带格式文字' },
  ];

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '单元格类型测试' },
    sharedStrings,
    sheets: [{
      name: 'CellTypes',
      cells,
      mergedCells: [],
      columnWidths: [120, 200, 150],
      rowHeights: [],
      hyperlinks: [],
    }],
  };
  await saveFixture('xlsx-cell-types', wb, 'xlsx');
}

async function xlsxMergedCells() {
  const sharedStrings: SharedStringEntry[] = [
    { text: '合并单元格测试' },
    { text: '跨行跨列' },
    { text: '数据' },
  ];

  const cells: Cell[][] = [
    [
      { value: '合并单元格测试', type: 'sharedString' },
      { value: null, type: 'string' },
      { value: null, type: 'string' },
    ],
    [
      { value: null, type: 'string' },
      { value: '跨行跨列', type: 'sharedString' },
      { value: null, type: 'string' },
    ],
    [
      { value: '数据', type: 'sharedString' },
      { value: null, type: 'string' },
      { value: '数据', type: 'sharedString' },
    ],
  ];

  const mergedCells: MergedCell[] = [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 2 },
    { startRow: 1, startCol: 1, endRow: 2, endCol: 2 },
  ];

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '合并单元格测试' },
    sharedStrings,
    sheets: [{
      name: 'MergedCells',
      cells,
      mergedCells,
      columnWidths: [150, 150, 150],
      rowHeights: [30, 30, 30],
      hyperlinks: [],
    }],
  };
  await saveFixture('xlsx-merged-cells', wb, 'xlsx');
}

async function xlsxHyperlinks() {
  const sharedStrings: SharedStringEntry[] = [
    { text: '超链接' },
    { text: '目标' },
    { text: '示例网站' },
    { text: 'https://example.com' },
    { text: '邮件' },
    { text: 'mailto:test@test.com' },
    { text: '文件链接' },
    { text: './other.xlsx' },
  ];

  const cells: Cell[][] = [
    [
      { value: '超链接', type: 'sharedString' },
      { value: '目标', type: 'sharedString' },
    ],
    [
      { value: '示例网站', type: 'sharedString' },
      { value: 'https://example.com', type: 'sharedString' },
    ],
    [
      { value: '邮件', type: 'sharedString' },
      { value: 'mailto:test@test.com', type: 'sharedString' },
    ],
    [
      { value: '文件链接', type: 'sharedString' },
      { value: './other.xlsx', type: 'sharedString' },
    ],
  ];

  const hyperlinks = [
    { ref: 'A2', url: 'https://example.com', tooltip: '访问示例' },
    { ref: 'B3', url: 'mailto:test@test.com' },
    { ref: 'A4', url: './other.xlsx' },
  ];

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '超链接测试' },
    sharedStrings,
    sheets: [{
      name: 'Hyperlinks',
      cells,
      mergedCells: [],
      columnWidths: [150, 250],
      rowHeights: [],
      hyperlinks,
    }],
  };
  await saveFixture('xlsx-hyperlinks', wb, 'xlsx');
}

async function xlsxAdvancedFeatures() {
  const sharedStrings: SharedStringEntry[] = [
    { text: '高级功能测试' },
    { text: 'A列' },
    { text: 'B列' },
    { text: 'C列' },
    { text: '数据1' },
    { text: '数据2' },
    { text: '数据3' },
    { text: '合计' },
    { text: '冻结窗格测试' },
  ];

  const cells: Cell[][] = [
    [
      { value: 'A列', type: 'sharedString' },
      { value: 'B列', type: 'sharedString' },
      { value: 'C列', type: 'sharedString' },
    ],
    [
      { value: '数据1', type: 'sharedString' },
      { value: 100, type: 'number' },
      { value: 'A', type: 'string' },
    ],
    [
      { value: '数据2', type: 'sharedString' },
      { value: 200, type: 'number' },
      { value: 'B', type: 'string' },
    ],
    [
      { value: '数据3', type: 'sharedString' },
      { value: 300, type: 'number' },
      { value: 'C', type: 'string' },
    ],
    [
      { value: '合计', type: 'sharedString' },
      { value: '', type: 'formula', formula: 'SUM(B2:B4)' },
      { value: null, type: 'string' },
    ],
  ];

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '高级功能测试' },
    sharedStrings,
    sheets: [{
      name: 'Advanced',
      cells,
      mergedCells: [],
      columnWidths: [120, 120, 120],
      rowHeights: [],
      hyperlinks: [],
      autoFilter: {
        ref: 'A1:C5',
        columns: [
          { colId: 0, filters: ['数据1', '数据2'] },
        ],
      },
      dataValidations: [{
        type: 'whole',
        operator: 'between',
        allowBlank: true,
        showErrorMessage: true,
        errorTitle: '输入错误',
        error: '请输入1-1000之间的整数',
        sqref: 'B2:B100',
        formula1: '1',
        formula2: '1000',
      }],
      conditionalFormats: [{
        sqref: 'B2:B4',
        rules: [
          { type: 'cellIs', priority: 1, formula: ['200'], dxfId: 0 },
        ],
      }],
      frozenPanes: { xSplit: 1, ySplit: 1, topLeftCell: 'B2' },
      zoomScale: 125,
      activeCell: 'B2',
      selections: [{ pane: 'topLeft', activeCell: 'B2', sqref: 'B2' }],
      defaultRowHeight: 20,
      defaultColWidth: 12,
      comments: [{
        ref: 'B2',
        authorId: 0,
        text: '这是一个批注',
      }],
      tables: [{
        id: 1,
        name: 'DataTable',
        displayName: 'DataTable',
        ref: 'A1:C5',
        headerRowCount: 1,
        columns: [
          { id: 1, name: 'A列' },
          { id: 2, name: 'B列' },
          { id: 3, name: 'C列' },
        ],
      }],
      tabColor: 'FF0000',
      state: 'visible',
      rowGroups: [
        { level: 1, collapsed: false },
      ],
      colGroups: [
        { level: 1 },
      ],
    }],
  };
  await saveFixture('xlsx-advanced-features', wb, 'xlsx');
}

async function xlsxMultipleSheets() {
  const sharedStrings: SharedStringEntry[] = [
    { text: 'Sheet1数据' },
    { text: 'Sheet2数据' },
    { text: 'Sheet3数据' },
  ];

  const makeSheet = (name: string, ssIdx: number, state?: 'visible' | 'hidden' | 'veryHidden'): Sheet => ({
    name,
    cells: [[{ value: sharedStrings[ssIdx].text, type: 'sharedString' }]],
    mergedCells: [],
    columnWidths: [],
    rowHeights: [],
    hyperlinks: [],
    state,
  });

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '多工作表测试' },
    sharedStrings,
    sheets: [
      makeSheet('可见表', 0, 'visible'),
      makeSheet('隐藏表', 1, 'hidden'),
      makeSheet('非常隐藏', 2, 'veryHidden'),
    ],
  };
  await saveFixture('xlsx-multiple-sheets', wb, 'xlsx');
}

async function xlsxProtection() {
  const sharedStrings: SharedStringEntry[] = [
    { text: '工作表保护测试' },
    { text: '受保护的数据' },
  ];

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '工作表保护测试' },
    sharedStrings,
    styles: {
      cellStyles: [{
        id: 'protected',
        protection: { locked: true, hidden: false },
      }],
      fonts: [{ name: 'Calibri', size: 11 }],
      fills: [],
      borders: [],
      numberFormats: [],
    },
    sheets: [{
      name: 'Protected',
      cells: [
        [{ value: '工作表保护测试', type: 'sharedString' }],
        [{ value: '受保护的数据', type: 'sharedString', style: 'protected' }],
      ],
      mergedCells: [],
      columnWidths: [200],
      rowHeights: [],
      hyperlinks: [],
      protection: {
        enabled: true,
        password: 'abc123',
        selectLockedCells: false,
        selectUnlockedCells: true,
        insertRows: false,
        deleteRows: false,
        formatCells: true,
        sort: false,
        autoFilter: false,
      },
    }],
  };
  await saveFixture('xlsx-protection', wb, 'xlsx');
}

async function xlsxPageSetup() {
  const sharedStrings: SharedStringEntry[] = [
    { text: '页面设置测试' },
  ];

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '页面设置测试' },
    sharedStrings,
    sheets: [{
      name: 'PageSetup',
      cells: [[{ value: '页面设置测试', type: 'sharedString' }]],
      mergedCells: [],
      columnWidths: [],
      rowHeights: [],
      hyperlinks: [],
      printArea: {
        sheet: true,
        fitToWidth: 1,
        fitToHeight: 0,
        pageMargins: { top: 1, right: 0.75, bottom: 1, left: 0.75, header: 0.5, footer: 0.5 },
      },
      pageSetup: {
        orientation: 'landscape',
        paperSize: 9,
        scale: 85,
        fitToWidth: 1,
        fitToHeight: 0,
        firstPageNumber: 1,
        useFirstPageNumber: true,
        horizontalDpi: 300,
        verticalDpi: 300,
      },
      headerFooter: {
        oddHeader: '&C&B页眉',
        oddFooter: '&L日期 &D&R第 &P 页，共 &N 页',
        evenHeader: '&C偶数页眉',
        evenFooter: '&L&R偶数页脚',
        firstHeader: '&C首页页眉',
        firstFooter: '&C首页页脚',
        differentFirst: true,
        differentOddEven: true,
      },
      printTitles: '$1:$1',
    }],
  };
  await saveFixture('xlsx-page-setup', wb, 'xlsx');
}

async function xlsxStyles() {
  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: { title: '样式测试' },
    sharedStrings: [{ text: '样式测试' }],
    styles: {
      cellStyles: [
        {
          id: 'header',
          font: { name: 'SimHei', size: 12, bold: true, color: 'FFFFFF' },
          fill: { patternType: 'solid', fgColor: '2F5496' },
          border: {
            top: { style: 'thin', color: '000000' },
            bottom: { style: 'thin', color: '000000' },
            left: { style: 'thin', color: '000000' },
            right: { style: 'thin', color: '000000' },
          },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        },
        {
          id: 'number',
          numberFormat: '#,##0.00',
          alignment: { horizontal: 'right' },
        },
        {
          id: 'percent',
          numberFormat: '0.00%',
        },
        {
          id: 'currency',
          font: { name: 'Calibri', size: 11 },
          numberFormat: '¥#,##0.00',
          alignment: { horizontal: 'right', indent: 1 },
        },
      ],
      fonts: [
        { name: 'SimHei', size: 12, bold: true, color: 'FFFFFF' },
        { name: 'Calibri', size: 11 },
      ],
      fills: [
        { patternType: 'solid', fgColor: '2F5496' },
      ],
      borders: [
        { top: { style: 'thin', color: '000000' }, bottom: { style: 'thin', color: '000000' }, left: { style: 'thin', color: '000000' }, right: { style: 'thin', color: '000000' } },
      ],
      numberFormats: [
        { id: '164', formatCode: '#,##0.00' },
        { id: '165', formatCode: '0.00%' },
        { id: '166', formatCode: '¥#,##0.00' },
      ],
    },
    sheets: [{
      name: 'Styles',
      cells: [
        [{ value: '样式测试', type: 'sharedString', style: 'header' }],
      ],
      mergedCells: [],
      columnWidths: [200],
      rowHeights: [],
      hyperlinks: [],
    }],
  };
  await saveFixture('xlsx-styles', wb, 'xlsx');
}

async function xlsxAllInOne() {
  const sharedStrings: SharedStringEntry[] = [
    { text: '综合测试' },
    { text: '字符串值' },
    { text: '合计' },
  ];

  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: FULL_CORE_META,
    appMeta: FULL_APP_META,
    customProperties: FULL_CUSTOM_PROPS,
    sharedStrings,
    styles: {
      cellStyles: [
        { id: 'default', font: { name: 'Calibri', size: 11 } },
      ],
      fonts: [{ name: 'Calibri', size: 11 }],
      fills: [],
      borders: [],
      numberFormats: [],
    },
    sheets: [
      {
        name: 'Main',
        cells: [
          [
            { value: '综合测试', type: 'sharedString', style: 'default' },
            { value: null, type: 'string' },
          ],
          [
            { value: '字符串值', type: 'sharedString' },
            { value: 100, type: 'number' },
          ],
          [
            { value: '合计', type: 'sharedString' },
            { value: '', type: 'formula', formula: 'B2*2' },
          ],
        ],
        mergedCells: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
        columnWidths: [150, 120],
        rowHeights: [],
        hyperlinks: [{ ref: 'A1', url: 'https://example.com' }],
        frozenPanes: { xSplit: 0, ySplit: 1 },
        state: 'visible',
        tabColor: '0070C0',
      },
      {
        name: 'Hidden',
        cells: [[{ value: '隐藏表', type: 'string' }]],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
        state: 'hidden',
      },
    ],
  };
  await saveFixture('xlsx-all-in-one', wb, 'xlsx');
}

async function xlsxEdgeCases() {
  const wb: XlsxWorkbook = {
    ...emptyXlsx(),
    meta: {},
    sharedStrings: [],
    sheets: [
      {
        name: 'Empty',
        cells: [],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
      },
      {
        name: 'Sparse',
        cells: [
          [],
          [],
          [{ value: 42, type: 'number' }],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
      },
      {
        name: 'SpecialChars',
        cells: [
          [{ value: '<html>&amp;"\'特殊', type: 'string' }],
          [{ value: '中文日本語한국어', type: 'string' }],
          [{ value: '\t\n\r', type: 'string' }],
        ],
        mergedCells: [],
        columnWidths: [],
        rowHeights: [],
        hyperlinks: [],
      },
    ],
  };
  await saveFixture('xlsx-edge-cases', wb, 'xlsx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// PPTX Fixtures
// ═══════════════════════════════════════════════════════════════════════════════

function defaultPptxScaffold(slides: Slide[], opts?: Partial<PptxPresentation>): PptxPresentation {
  const layout: SlideLayout = {
    id: '1',
    name: 'Blank',
    type: 'obj',
    placeholders: [],
  };
  const master: SlideMaster = {
    id: '1',
    layouts: [layout],
    background: { type: 'solid', color: 'FFFFFF' },
    txStyles: {
      titleStyle: { levels: [{ level: 1, alignment: 'ctr' }] },
      bodyStyle: { levels: [{ level: 1 }] },
      otherStyle: { levels: [{ level: 1 }] },
    },
  };
  return {
    meta: {},
    slides,
    masters: [master],
    layouts: [layout],
    theme: {
      colorScheme: {
        name: 'Office',
        colors: {
          dk1: '1F3864', lt1: 'FFFFFF', dk2: '4472C4', lt2: 'E7E6E6',
          accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
          accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
          hlink: '0563C1', folHlink: '954F72',
        },
      },
      fontScheme: { name: 'Office', majorFont: 'Calibri', minorFont: 'Calibri' },
    },
    slideSize: { width: 9144000, height: 6858000 },
    notesSize: { width: 6858000, height: 9144000 },
    ...opts,
  };
}

async function pptxMetaFull() {
  const pres = defaultPptxScaffold(
    [{ elements: [{ type: 'text', content: '元数据测试', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('元数据测试幻灯片')] }] }],
    { meta: FULL_CORE_META, appMeta: FULL_APP_META, customProperties: FULL_CUSTOM_PROPS },
  );
  await saveFixture('pptx-meta-full', pres, 'pptx');
}

async function pptxContentTextShapes() {
  const elements: SlideElement[] = [
    {
      type: 'text',
      content: '标题',
      position: { x: 457200, y: 274320, width: 8229600, height: 1143000 },
      paragraphs: [
        { type: 'paragraph', runs: [{ text: '大标题', bold: true, fontSize: 36, color: '2F5496' }] },
      ],
      style: {
        fill: { color: 'F2F2F2', type: 'solid' },
        border: { color: '2F5496', width: 12700, style: 'solid' },
      },
      rotation: 0,
      presetGeom: 'rect',
      bodyProperties: {
        anchor: 'ctr',
        wrap: 'square',
        leftInset: 91440,
        topInset: 45720,
        rightInset: 91440,
        bottomInset: 45720,
      },
    },
    {
      type: 'text',
      content: '正文',
      position: { x: 457200, y: 1600200, width: 8229600, height: 4572000 },
      paragraphs: [
        {
          type: 'paragraph',
          runs: [
            { text: '普通文本 ', fontSize: 18 },
            { text: '粗体', bold: true, fontSize: 18 },
            { text: ' ' },
            { text: '斜体', italic: true, fontSize: 18 },
            { text: ' ' },
            { text: '下划线', underline: true, fontSize: 18 },
          ],
        },
        {
          type: 'paragraph',
          runs: [
            { text: '红色文字 ', color: 'FF0000', fontSize: 18 },
            { text: '蓝色大字', color: '0000FF', fontSize: 24 },
          ],
          properties: { alignment: 'center' },
        },
        {
          type: 'paragraph',
          runs: [
            { text: '删除线', strike: true, fontSize: 18 },
            { text: ' ' },
            { text: '上标', superscript: true, fontSize: 14 },
            { text: ' ' },
            { text: '下标', subscript: true, fontSize: 14 },
          ],
        },
      ],
      bodyProperties: {
        anchor: 't',
        wrap: 'square',
        autoFit: 'normal',
      },
      listStyle: {
        defaultParagraphProperties: [
          { level: 1, alignment: 'l', indent: -228600, marL: 342900 },
        ],
      },
    },
    {
      type: 'text',
      content: '带占位符',
      position: { x: 0, y: 0, width: 9144000, height: 685800 },
      paragraphs: [p('占位符文本')],
      placeholder: { type: 'title', index: 0 },
    },
  ];

  const pres = defaultPptxScaffold([{ elements }]);
  await saveFixture('pptx-content-text-shapes', pres, 'pptx');
}

async function pptxContentShapes() {
  const elements: SlideElement[] = [
    {
      type: 'text',
      content: '矩形',
      position: { x: 457200, y: 457200, width: 2743200, height: 1828800 },
      paragraphs: [p('矩形')],
      presetGeom: 'rect',
      style: {
        fill: { type: 'solid', color: '4472C4' },
        border: { color: '2F5496', width: 25400, style: 'solid' },
        shadow: { type: 'outer', color: '808080', blur: 50800, offsetX: 25400, offsetY: 25400 },
        opacity: 90,
      },
    },
    {
      type: 'text',
      content: '圆形',
      position: { x: 3657600, y: 457200, width: 1828800, height: 1828800 },
      paragraphs: [p('圆')],
      presetGeom: 'ellipse',
      style: {
        fill: { type: 'gradient', gradientFill: { type: 'linear', angle: 45, stops: [{ position: 0, color: 'FF0000' }, { position: 100, color: '0000FF' }] } },
        border: { color: '333333', width: 12700, style: 'dashed', dashType: 'dash' },
      },
    },
    {
      type: 'text',
      content: '带效果',
      position: { x: 5943600, y: 457200, width: 2743200, height: 1828800 },
      paragraphs: [p('效果')],
      presetGeom: 'roundRect',
      style: {
        fill: { type: 'solid', color: '70AD47' },
        glow: { color: 'FFD700', radius: 50800 },
        softEdge: { radius: 25400 },
      },
    },
  ];

  const pres = defaultPptxScaffold([{ elements }]);
  await saveFixture('pptx-content-shapes', pres, 'pptx');
}

async function pptxContentTables() {
  const elements: SlideElement[] = [
    {
      type: 'text',
      content: '表格标题',
      position: { x: 457200, y: 274320, width: 8229600, height: 457200 },
      paragraphs: [p('表格示例')],
    },
    {
      type: 'table',
      position: { x: 457200, y: 914400, width: 8229600, height: 2743200 },
      rows: [
        {
          cells: [
            { content: [p('列A')] },
            { content: [p('列B')] },
            { content: [p('列C')] },
          ],
          height: 457200,
        },
        {
          cells: [
            { content: [p('1')] },
            { content: [p('2')] },
            { content: [p('3')] },
          ],
          height: 457200,
        },
        {
          cells: [
            { content: [p('X')] },
            { content: [{ type: 'paragraph', runs: [{ text: '合并', bold: true }] }] },
            { content: [p('Z')] },
          ],
          height: 457200,
        },
      ],
      style: {
        fill: { type: 'solid', color: '2F5496' },
      },
    },
  ];

  const pres = defaultPptxScaffold([{ elements }]);
  await saveFixture('pptx-content-tables', pres, 'pptx');
}

async function pptxContentGroups() {
  const elements: SlideElement[] = [
    {
      type: 'group',
      position: { x: 457200, y: 457200, width: 4572000, height: 3657600 },
      children: [
        {
          type: 'text',
          content: '组内文本',
          position: { x: 0, y: 0, width: 2286000, height: 914400 },
          paragraphs: [p('在组合中')],
          style: { fill: { type: 'solid', color: 'E2EFDA' } },
        },
        {
          type: 'text',
          content: '组内矩形',
          position: { x: 2286000, y: 0, width: 2286000, height: 914400 },
          paragraphs: [p('也是组合')],
          presetGeom: 'rect',
          style: { fill: { type: 'solid', color: 'D6E4F0' } },
        },
      ],
      childOffset: { x: 0, y: 0 },
      childExtent: { width: 4572000, height: 3657600 },
    },
  ];

  const pres = defaultPptxScaffold([{ elements }]);
  await saveFixture('pptx-content-groups', pres, 'pptx');
}

async function pptxContentTransitions() {
  const slides: Slide[] = [
    {
      elements: [{ type: 'text', content: '淡入', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('淡入切换')] }],
      transition: { type: 'fade', duration: 1500, advClick: true },
    },
    {
      elements: [{ type: 'text', content: '推进', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('推进切换')] }],
      transition: { type: 'push', duration: 800, direction: 'l', advClick: true, advTime: 5000 },
    },
    {
      elements: [{ type: 'text', content: '无切换', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('无切换效果')] }],
    },
  ];

  const pres = defaultPptxScaffold(slides);
  await saveFixture('pptx-content-transitions', pres, 'pptx');
}

async function pptxContentAnimations() {
  const slides: Slide[] = [
    {
      elements: [
        {
          type: 'text',
          content: '动画对象',
          position: { x: 457200, y: 457200, width: 4572000, height: 2286000 },
          paragraphs: [p('点击后淡入')],
        },
      ],
      animations: [
        {
          trigger: 'onClick',
          type: 'fade',
          shapeId: '2',
          duration: 1000,
          delay: 0,
        },
        {
          trigger: 'withPrevious',
          type: 'flyIn',
          shapeId: '2',
          duration: 500,
          delay: 200,
          direction: 'fromLeft',
        },
      ],
    },
  ];

  const pres = defaultPptxScaffold(slides);
  await saveFixture('pptx-content-animations', pres, 'pptx');
}

async function pptxContentNotes() {
  const slides: Slide[] = [
    {
      elements: [{ type: 'text', content: '有备注', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('第一张幻灯片')] }],
      notes: '这是第一张幻灯片的备注内容，包含演讲者提示。',
    },
    {
      elements: [{ type: 'text', content: '无备注', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('第二张幻灯片')] }],
    },
  ];

  const pres = defaultPptxScaffold(slides);
  await saveFixture('pptx-content-notes', pres, 'pptx');
}

async function pptxContentMasterLayout() {
  const layout1: SlideLayout = {
    id: '1',
    name: 'Title Slide',
    type: 'title',
    placeholders: [
      { type: 'ctrTitle', position: { x: 457200, y: 2130425, width: 8229600, height: 1600200 } },
      { type: 'subTitle', position: { x: 457200, y: 3886200, width: 8229600, height: 1219200 } },
    ],
  };
  const layout2: SlideLayout = {
    id: '2',
    name: 'Title and Content',
    type: 'obj',
    placeholders: [
      { type: 'title', position: { x: 457200, y: 274320, width: 8229600, height: 1143000 }, index: 0 },
      { type: 'body', position: { x: 457200, y: 1600200, width: 8229600, height: 4572000 }, index: 1 },
    ],
  };
  const layout3: SlideLayout = {
    id: '3',
    name: 'Blank',
    type: 'obj',
    placeholders: [],
  };

  const master: SlideMaster = {
    id: '1',
    layouts: [layout1, layout2, layout3],
    background: { type: 'solid', color: 'FFFFFF' },
    txStyles: {
      titleStyle: {
        levels: [
          { level: 1, alignment: 'ctr', fontScale: 100, spcBef: 0, spcAft: 0 },
        ],
      },
      bodyStyle: {
        levels: [
          { level: 1, alignment: 'l', marL: 0, indent: 0, fontScale: 80 },
          { level: 2, alignment: 'l', marL: 228600, indent: -228600, fontScale: 70 },
        ],
      },
      otherStyle: {
        levels: [{ level: 1 }],
      },
    },
  };

  const slides: Slide[] = [
    {
      elements: [
        {
          type: 'text',
          content: '标题',
          position: { x: 457200, y: 2130425, width: 8229600, height: 1600200 },
          paragraphs: [{ type: 'paragraph', runs: [{ text: '标题页', bold: true, fontSize: 44 }], properties: { alignment: 'center' } }],
          placeholder: { type: 'ctrTitle' },
        },
        {
          type: 'text',
          content: '副标题',
          position: { x: 457200, y: 3886200, width: 8229600, height: 1219200 },
          paragraphs: [{ type: 'paragraph', runs: [{ text: '副标题文本', fontSize: 24 }], properties: { alignment: 'center' } }],
          placeholder: { type: 'subTitle' },
        },
      ],
      layout: '1',
    },
    {
      elements: [
        {
          type: 'text',
          content: '内容标题',
          position: { x: 457200, y: 274320, width: 8229600, height: 1143000 },
          paragraphs: [{ type: 'paragraph', runs: [{ text: '内容页', bold: true, fontSize: 32 }] }],
          placeholder: { type: 'title', index: 0 },
        },
        {
          type: 'text',
          content: '内容',
          position: { x: 457200, y: 1600200, width: 8229600, height: 4572000 },
          paragraphs: [
            { type: 'paragraph', runs: [{ text: '要点一', fontSize: 18 }] },
            { type: 'paragraph', runs: [{ text: '要点二', fontSize: 18 }] },
          ],
          placeholder: { type: 'body', index: 1 },
        },
      ],
      layout: '2',
    },
    {
      elements: [
        {
          type: 'text',
          content: '空白页',
          position: { x: 457200, y: 2743200, width: 8229600, height: 1143000 },
          paragraphs: [p('空白布局上的内容')],
        },
      ],
      layout: '3',
    },
  ];

  const pres: PptxPresentation = {
    meta: { title: '母版布局测试' },
    slides,
    masters: [master],
    layouts: [layout1, layout2, layout3],
    theme: {
      colorScheme: {
        name: 'Office',
        colors: {
          dk1: '1F3864', lt1: 'FFFFFF', dk2: '4472C4', lt2: 'E7E6E6',
          accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
          accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
          hlink: '0563C1', folHlink: '954F72',
        },
      },
      fontScheme: { name: 'Office', majorFont: 'Calibri', minorFont: 'Calibri' },
    },
    slideSize: { width: 9144000, height: 6858000 },
    notesSize: { width: 6858000, height: 9144000 },
  };
  await saveFixture('pptx-content-master-layout', pres, 'pptx');
}

async function pptxContentTheme() {
  const slides: Slide[] = [
    {
      elements: [{ type: 'text', content: '主题测试', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('应用了自定义主题')] }],
      background: { type: 'solid', color: 'F5F5F5' },
    },
  ];

  const pres = defaultPptxScaffold(slides, {
    meta: { title: '主题测试' },
    theme: {
      colorScheme: {
        name: 'Custom',
        colors: {
          dk1: '1F3864',
          lt1: 'FFFFFF',
          dk2: '4472C4',
          lt2: 'E7E6E6',
          accent1: '4472C4',
          accent2: 'ED7D31',
          accent3: 'A5A5A5',
          accent4: 'FFC000',
          accent5: '5B9BD5',
          accent6: '70AD47',
          hlink: '0563C1',
          folHlink: '954F72',
        },
      },
      fontScheme: {
        name: 'Custom',
        majorFont: '微软雅黑',
        minorFont: '宋体',
      },
      formatScheme: {
        fillStyles: [
          { type: 'solid', color: 'FFFFFF' },
          { type: 'solid', color: 'E7E6E6' },
          { type: 'solid', color: '4472C4' },
        ],
        lineStyles: [
          { color: '4472C4', width: 25400, style: 'solid' },
          { color: 'A5A5A5', width: 12700, style: 'solid' },
          { color: 'D9D9D9', width: 6350, style: 'dashed', dashType: 'dash' },
        ],
        effectStyles: [
          { shadow: { type: 'outer', color: '808080', blur: 50800, offsetX: 25400, offsetY: 25400 } },
          { glow: { color: '4472C4', radius: 25400 } },
          { softEdge: { radius: 12700 } },
        ],
        bgFillStyles: [
          { type: 'solid', color: 'FFFFFF' },
          { type: 'solid', color: 'F5F5F5' },
          { type: 'solid', color: 'E7E6E6' },
        ],
      },
    },
  });
  await saveFixture('pptx-content-theme', pres, 'pptx');
}

async function pptxContentBackgrounds() {
  const slides: Slide[] = [
    {
      elements: [{ type: 'text', content: '纯色背景', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('纯色背景')] }],
      background: { type: 'solid', color: 'E2EFDA' },
      showMasterSp: false,
      showMasterPhAnim: false,
    },
    {
      elements: [{ type: 'text', content: '渐变背景', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('渐变背景')] }],
      background: {
        type: 'gradient',
        gradientFill: {
          type: 'linear',
          angle: 270,
          stops: [
            { position: 0, color: 'FFFFFF' },
            { position: 100, color: '4472C4' },
          ],
        },
      },
    },
    {
      elements: [{ type: 'text', content: '默认背景', position: { x: 0, y: 0, width: 9144000, height: 6858000 }, paragraphs: [p('使用母版背景')] }],
      showMasterSp: true,
    },
  ];

  const pres = defaultPptxScaffold(slides, { meta: { title: '背景测试' } });
  await saveFixture('pptx-content-backgrounds', pres, 'pptx');
}

async function pptxAllInOne() {
  const layout1: SlideLayout = {
    id: '1',
    name: 'Title Slide',
    type: 'title',
    placeholders: [
      { type: 'ctrTitle', position: { x: 457200, y: 2130425, width: 8229600, height: 1600200 } },
    ],
  };
  const layout2: SlideLayout = {
    id: '2',
    name: 'Content',
    type: 'obj',
    placeholders: [
      { type: 'title', position: { x: 457200, y: 274320, width: 8229600, height: 1143000 }, index: 0 },
    ],
  };

  const master: SlideMaster = {
    id: '1',
    layouts: [layout1, layout2],
    background: { type: 'solid', color: 'FFFFFF' },
    txStyles: {
      titleStyle: { levels: [{ level: 1, alignment: 'ctr' }] },
      bodyStyle: { levels: [{ level: 1 }] },
      otherStyle: { levels: [{ level: 1 }] },
    },
  };

  const slides: Slide[] = [
    {
      elements: [
        {
          type: 'text',
          content: '综合测试标题',
          position: { x: 457200, y: 2130425, width: 8229600, height: 1600200 },
          paragraphs: [{ type: 'paragraph', runs: [{ text: '综合测试', bold: true, fontSize: 44, color: '2F5496' }], properties: { alignment: 'center' } }],
          placeholder: { type: 'ctrTitle' },
        },
      ],
      transition: { type: 'fade', duration: 1000 },
      layout: '1',
      background: { type: 'solid', color: 'F2F2F2' },
    },
    {
      elements: [
        {
          type: 'text',
          content: '内容页标题',
          position: { x: 457200, y: 274320, width: 8229600, height: 1143000 },
          paragraphs: [p('功能概览')],
          placeholder: { type: 'title', index: 0 },
        },
        {
          type: 'text',
          content: '文本框',
          position: { x: 457200, y: 1600200, width: 4114800, height: 2286000 },
          paragraphs: [
            { type: 'paragraph', runs: [{ text: '粗体斜体', bold: true, italic: true, fontSize: 18 }] },
            { type: 'paragraph', runs: [{ text: '彩色文字', color: 'FF0000', fontSize: 18 }] },
          ],
          style: {
            fill: { type: 'solid', color: 'E2EFDA' },
            border: { color: '70AD47', width: 12700, style: 'solid' },
          },
          bodyProperties: { anchor: 't', wrap: 'square' },
        },
        {
          type: 'text',
          content: '带效果',
          position: { x: 5029200, y: 1600200, width: 3657600, height: 2286000 },
          paragraphs: [p('阴影和发光效果')],
          presetGeom: 'roundRect',
          style: {
            fill: { type: 'solid', color: 'D6E4F0' },
            shadow: { type: 'outer', color: '808080', blur: 50800, offsetX: 25400, offsetY: 25400 },
            glow: { color: '4472C4', radius: 25400 },
          },
        },
      ],
      transition: { type: 'push', direction: 'l', duration: 500 },
      animations: [
        { trigger: 'onClick', type: 'fade', shapeId: '2', duration: 800 },
      ],
      notes: '这是第二张幻灯片的备注。',
      layout: '2',
      showMasterSp: true,
      showMasterPhAnim: true,
    },
    {
      elements: [
        {
          type: 'table',
          position: { x: 457200, y: 457200, width: 8229600, height: 2743200 },
          rows: [
            { cells: [{ content: [p('功能')] }, { content: [p('状态')] }], height: 457200 },
            { cells: [{ content: [p('文本')] }, { content: [p('OK')] }], height: 457200 },
            { cells: [{ content: [p('表格')] }, { content: [p('OK')] }], height: 457200 },
            { cells: [{ content: [p('动画')] }, { content: [p('OK')] }], height: 457200 },
          ],
        },
        {
          type: 'group',
          position: { x: 457200, y: 3657600, width: 8229600, height: 2286000 },
          children: [
            {
              type: 'text',
              content: '组1',
              position: { x: 0, y: 0, width: 2743200, height: 1143000 },
              paragraphs: [p('组合元素A')],
              style: { fill: { type: 'solid', color: 'FFF2CC' } },
            },
            {
              type: 'text',
              content: '组2',
              position: { x: 2743200, y: 0, width: 2743200, height: 1143000 },
              paragraphs: [p('组合元素B')],
              style: { fill: { type: 'solid', color: 'D6E4F0' } },
            },
          ],
        },
      ],
      layout: '2',
    },
  ];

  const pres: PptxPresentation = {
    meta: FULL_CORE_META,
    appMeta: FULL_APP_META,
    customProperties: FULL_CUSTOM_PROPS,
    slides,
    masters: [master],
    layouts: [layout1, layout2],
    theme: {
      colorScheme: {
        name: 'Office',
        colors: {
          dk1: '1F3864', lt1: 'FFFFFF', dk2: '4472C4', lt2: 'E7E6E6',
          accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
          accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
          hlink: '0563C1', folHlink: '954F72',
        },
      },
      fontScheme: { name: 'Office', majorFont: 'Calibri', minorFont: 'Calibri' },
    },
    slideSize: { width: 9144000, height: 6858000 },
    notesSize: { width: 6858000, height: 9144000 },
  };
  await saveFixture('pptx-all-in-one', pres, 'pptx');
}

async function pptxEdgeCases() {
  const slides: Slide[] = [
    {
      elements: [
        {
          type: 'text',
          content: '',
          position: { x: 0, y: 0, width: 9144000, height: 6858000 },
          paragraphs: [],
        },
      ],
    },
    {
      elements: [
        {
          type: 'text',
          content: '特殊字符',
          position: { x: 0, y: 0, width: 9144000, height: 6858000 },
          paragraphs: [
            { type: 'paragraph', runs: [{ text: '<html>&amp;"\'中文日本語한국어' }] },
            { type: 'paragraph', runs: [{ text: '' }] },
            { type: 'paragraph', runs: [] },
          ],
        },
      ],
    },
  ];

  const pres = defaultPptxScaffold(slides, { meta: {} });
  await saveFixture('pptx-edge-cases', pres, 'pptx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// 主函数
// ═══════════════════════════════════════════════════════════════════════════════

// ─── 异常数据校验 ──────────────────────────────────────────────────────────────

function assertThrows(label: string, fn: () => void) {
  try {
    fn();
    console.log(`  ✗ ${label} — 未抛出异常`);
  } catch (e: any) {
    console.log(`  ✓ ${label} — 捕获校验异常`);
  }
}

async function invalidDocx() {
  // 1. commentId 引用不存在的 comment
  const badCommentRef: DocxDocument = {
    ...emptyDocx(),
    comments: [{ id: '1', author: 'A', date: '2024-01-01T00:00:00Z', content: [p('批注')] }],
    body: { blocks: [p('文本', { commentId: '999' })] },
  };
  assertThrows('docx: commentId 引用不存在', () => {
    const issues = validateDocx(badCommentRef);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 2. bookmark 未成对
  const badBookmark: DocxDocument = {
    ...emptyDocx(),
    body: {
      blocks: [
        { type: 'bookmarkStart', id: '1', name: 'bm1' },
        p('没有 bookmarkEnd'),
      ],
    },
  };
  assertThrows('docx: bookmark 未成对', () => {
    const issues = validateDocx(badBookmark);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 3. hyperlink 缺少 relationshipId
  const badHyperlink: DocxDocument = {
    ...emptyDocx(),
    body: {
      blocks: [
        { type: 'hyperlink', relationshipId: '', url: 'https://example.com', runs: [{ text: '链接' }] },
      ],
    },
  };
  assertThrows('docx: hyperlink 缺少 relationshipId', () => {
    const issues = validateDocx(badHyperlink);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 4. numbering numId 引用不存在
  const badNumbering: DocxDocument = {
    ...emptyDocx(),
    numbering: { abstractNums: [], nums: [] },
    body: {
      blocks: [
        { type: 'paragraph', runs: [{ text: '列表项' }], numbering: { level: 0, numId: '999' } },
      ],
    },
  };
  assertThrows('docx: numId 引用不存在', () => {
    const issues = validateDocx(badNumbering);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 5. body.blocks 为空
  const emptyBody: DocxDocument = { ...emptyDocx(), body: { blocks: [] } };
  assertThrows('docx: body.blocks 为空', () => {
    const issues = validateDocx(emptyBody);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });
}

async function invalidXlsx() {
  // 1. sheets 为空
  const emptySheets: XlsxWorkbook = { ...emptyXlsx(), sheets: [] };
  assertThrows('xlsx: sheets 为空', () => {
    const issues = validateXlsx(emptySheets);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 2. sharedString 引用不存在
  const badSS: XlsxWorkbook = {
    ...emptyXlsx(),
    sharedStrings: [{ text: '存在的' }],
    sheets: [{
      name: 'Sheet1',
      cells: [[{ value: '不存在的', type: 'sharedString' }]],
      mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [],
    }],
  };
  assertThrows('xlsx: sharedString 引用不存在', () => {
    const issues = validateXlsx(badSS);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 3. mergedCells 范围非法
  const badMerge: XlsxWorkbook = {
    ...emptyXlsx(),
    sheets: [{
      name: 'Sheet1',
      cells: [[{ value: 'A', type: 'string' }]],
      mergedCells: [{ startRow: 5, startCol: 5, endRow: 1, endCol: 1 }],
      columnWidths: [], rowHeights: [], hyperlinks: [],
    }],
  };
  assertThrows('xlsx: mergedCells 范围非法', () => {
    const issues = validateXlsx(badMerge);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 4. sheet name 为空
  const badName: XlsxWorkbook = {
    ...emptyXlsx(),
    sheets: [{
      name: '',
      cells: [],
      mergedCells: [], columnWidths: [], rowHeights: [], hyperlinks: [],
    }],
  };
  assertThrows('xlsx: sheet name 为空', () => {
    const issues = validateXlsx(badName);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });
}

async function invalidPptx() {
  const sampleSlide: Slide = {
    elements: [{ type: 'text', content: 'hi', position: { x: 0, y: 0, width: 100, height: 100 }, paragraphs: [p('hi')] }],
  };

  // 1. 缺少 theme
  const noTheme: PptxPresentation = {
    ...defaultPptxScaffold([sampleSlide]),
    theme: undefined as any,
  };
  assertThrows('pptx: 缺少 theme', () => {
    const issues = validatePptx(noTheme);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 2. masters 为空
  const noMasters: PptxPresentation = {
    ...defaultPptxScaffold([sampleSlide]),
    masters: [],
  };
  assertThrows('pptx: masters 为空', () => {
    const issues = validatePptx(noMasters);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 3. image 缺少 relationshipId
  const badImage: PptxPresentation = defaultPptxScaffold([{
    elements: [{ type: 'image', relationshipId: '', position: { x: 0, y: 0, width: 100, height: 100 } }],
  }]);
  assertThrows('pptx: image 缺少 relationshipId', () => {
    const issues = validatePptx(badImage);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });

  // 4. table rows 为空
  const emptyTable: PptxPresentation = defaultPptxScaffold([{
    elements: [{ type: 'table', position: { x: 0, y: 0, width: 100, height: 100 }, rows: [] }],
  }]);
  assertThrows('pptx: table rows 为空', () => {
    const issues = validatePptx(emptyTable);
    const errors = issues.filter(i => i.level === 'error');
    if (errors.length > 0) throw new Error(formatValidationReport(errors));
  });
}

async function main() {
  console.log('='.repeat(60));
  console.log('Office Meta Parser - Test Fixture Generator');
  console.log('='.repeat(60));
  console.log(`Output: ${OUTPUT_DIR}\n`);

  // DOCX
  console.log('DOCX fixtures:');
  await docxMetaFull();
  await docxMetaCoreOnly();
  await docxMetaAppOnly();
  await docxMetaCustomOnly();
  await docxContentParagraphs();
  await docxContentTables();
  await docxContentHyperlinks();
  await docxContentBookmarks();
  await docxContentCommentsRevisions();
  await docxContentHeadersFooters();
  await docxContentNumbering();
  await docxContentFootnotesEndnotes();
  await docxContentStyles();
  await docxContentSettingsFontsTheme();
  await docxContentFields();
  await docxAllInOne();
  await docxEdgeCases();

  // XLSX
  console.log('\nXLSX fixtures:');
  await xlsxMetaFull();
  await xlsxCellTypes();
  await xlsxMergedCells();
  await xlsxHyperlinks();
  await xlsxAdvancedFeatures();
  await xlsxMultipleSheets();
  await xlsxProtection();
  await xlsxPageSetup();
  await xlsxStyles();
  await xlsxAllInOne();
  await xlsxEdgeCases();

  // PPTX
  console.log('\nPPTX fixtures:');
  await pptxMetaFull();
  await pptxContentTextShapes();
  await pptxContentShapes();
  await pptxContentTables();
  await pptxContentGroups();
  await pptxContentTransitions();
  await pptxContentAnimations();
  await pptxContentNotes();
  await pptxContentMasterLayout();
  await pptxContentTheme();
  await pptxContentBackgrounds();
  await pptxAllInOne();
  await pptxEdgeCases();

  // 异常数据校验
  console.log('\nValidation (expected errors):');
  await invalidDocx();
  await invalidXlsx();
  await invalidPptx();

  console.log('\n' + '='.repeat(60));
  console.log('Done!');
  console.log('='.repeat(60));
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
