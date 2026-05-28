# Office Meta Parser 设计文档

## 概述

构建纯 TypeScript npm 项目，实现 Office（PPTX、DOCX、XLSX）元数据解析与序列化。

提供两种 JSON 视图：
- **原始视图**：1:1 对应 Office XML 结构，用于调试和保真往返转换
- **语义视图**：语义化抽象结构，用于业务层直接使用

支持两种输入模式：
- 二进制文件（.pptx/.docx/.xlsx ArrayBuffer）
- XML 片段（已有 XML 字符串）

## 架构方案

采用 **Monorepo + 子路径导出** 方案，单一 npm 包，按需导入。

### 目录结构

```
office-meta/
├── src/
│   ├── core/                  # 共享底层
│   │   ├── zip.ts             # ZIP 解压/打包（JSZip 封装）
│   │   ├── xml.ts             # XML 解析/序列化（fast-xml-parser 封装）
│   │   ├── rels.ts            # .rels 关系文件解析
│   │   ├── content-type.ts    # [Content_Types].xml 处理
│   │   └── types.ts           # 公共类型定义
│   │
│   ├── docx/                  # Word 编解码
│   │   ├── parser.ts          # XML -> Raw JSON
│   │   ├── semantic.ts        # Raw JSON -> Semantic JSON
│   │   ├── serializer.ts      # JSON -> XML
│   │   ├── types.ts           # DOCX 专用类型
│   │   └── index.ts           # 子模块入口
│   │
│   ├── xlsx/                  # Excel 编解码（结构同上）
│   ├── pptx/                  # PPT 编解码（结构同上）
│   │
│   └── index.ts               # 主入口（re-export 所有子模块）
│
├── package.json
├── tsconfig.json
└── tsup.config.ts
```

### 子路径导出

```json
{
  "exports": {
    ".": "./dist/index.js",
    "./core": "./dist/core/index.js",
    "./docx": "./dist/docx/index.js",
    "./xlsx": "./dist/xlsx/index.js",
    "./pptx": "./dist/pptx/index.js"
  }
}
```

## Core 模块设计

Core 模块封装三种格式共享的底层能力，不包含格式特定逻辑。

### zip.ts — ZIP 操作封装

```typescript
interface ZipEntry {
  path: string;           // e.g. "word/document.xml"
  data: ArrayBuffer;
}

async function unzip(buffer: ArrayBuffer): Promise<ZipEntry[]>;
async function zip(entries: ZipEntry[]): Promise<ArrayBuffer>;
```

### xml.ts — XML 操作封装

```typescript
interface ParsedNode {
  tag: string;
  attrs: Record<string, string>;
  children: (ParsedNode | string)[];
}

function parseXml(xml: string): ParsedNode;
function serializeXml(node: ParsedNode): string;
```

### rels.ts — 关系文件解析

```typescript
interface Relationship {
  id: string;
  type: string;        // e.g. "http://.../image"
  target: string;      // e.g. "media/image1.png"
  targetMode?: string; // "External" | undefined
}

function parseRels(xml: string): Relationship[];
function serializeRels(rels: Relationship[]): string;
```

### content-type.ts — ContentTypes 处理

```typescript
interface ContentType {
  partName: string;
  contentType: string;
}

function parseContentTypes(xml: string): ContentType[];
function serializeContentTypes(types: ContentType[]): string;
```

### types.ts — 公共类型

```typescript
interface ParseResult<TSemantic> {
  raw: RawDocument;
  semantic: TSemantic;
}

interface RawDocument {
  entries: ZipEntry[];
  rels: Map<string, Relationship[]>;
  contentTypes: ContentType[];
  parts: Map<string, ParsedNode>;
}
```

## 语义化 JSON Schema

### DOCX 语义模型

```typescript
interface DocxDocument {
  meta: DocumentMeta;
  styles: StyleDefinitions;
  body: DocxBody;
  comments?: Comment[];
  trackChanges?: Revision[];
}

interface DocumentMeta {
  title?: string;
  author?: string;
  created?: string;
  modified?: string;
  lastModifiedBy?: string;
}

interface StyleDefinitions {
  paragraphStyles: ParagraphStyle[];
  characterStyles: CharacterStyle[];
  tableStyles: TableStyle[];
}

interface DocxBody {
  blocks: DocxBlock[];
}

type DocxBlock = Paragraph | Table | Image;

interface Paragraph {
  type: 'paragraph';
  style?: string;
  runs: TextRun[];
  properties?: ParagraphProperties;
}

interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;       // 单位：半磅
  color?: string;
  fontFamily?: string;
}
```

### XLSX 语义模型

```typescript
interface XlsxWorkbook {
  meta: DocumentMeta;
  sheets: Sheet[];
  styles: CellStyleDefinitions;
  sharedStrings: string[];
}

interface Sheet {
  name: string;
  cells: Cell[][];
  mergedCells: MergedCell[];
  columnWidths: number[];
  rowHeights: number[];
}

interface Cell {
  value: string | number | boolean | null;
  formula?: string;
  style?: CellStyleRef;
  type: 'string' | 'number' | 'boolean' | 'formula' | 'sharedString';
}
```

### PPTX 语义模型

```typescript
interface PptxPresentation {
  meta: DocumentMeta;
  slides: Slide[];
  masters: SlideMaster[];
  layouts: SlideLayout[];
  theme?: Theme;
}

interface Slide {
  elements: SlideElement[];
  transition?: Transition;
  notes?: string;
}

type SlideElement = TextShape | ImageShape | GroupShape;

interface TextShape {
  type: 'text';
  content: string;
  position: Position;
  style?: ShapeStyle;
  paragraphs: Paragraph[];
}
```

## 编解码流程

### 解析流程（Decoder）

```
二进制文件 (ArrayBuffer)
  │
  ├─ unzip() ──> ZipEntry[]
  │
  ├─ parseContentTypes() ──> content types
  ├─ parseRels() ──> 关系映射
  ├─ parseXml() 各个 XML part ──> ParsedNode map
  │
  ├─ RawDocument (原始视图，直接返回)
  │
  └─ Semantic Converter (格式特定)
       ├─ 提取 meta (core.xml / app.xml)
       ├─ 解析 styles (styles.xml / theme)
       ├─ 解析 body (document.xml / workbook.xml / presentation.xml)
       └─ DocxDocument / XlsxWorkbook / PptxPresentation (语义视图)
```

### 序列化流程（Encoder）

```
Semantic JSON 或 Raw JSON
  │
  ├─ 如果是语义视图：先转换为 Raw 视图
  │
  ├─ serializeXml() 各个 part ──> XML strings
  ├─ serializeRels() ──> .rels XML
  ├─ serializeContentTypes() ──> [Content_Types].xml
  │
  ├─ 组装 ZipEntry[]
  └─ zip() ──> ArrayBuffer
```

### 设计决策

- 语义 -> 原始 的转换是单向的，语义视图不保留所有原始信息（如未知的自定义属性）
- 100% 保真还原应使用原始视图进行往返转换
- 语义视图适合业务创建新文档或做有损但语义清晰的编辑

## 对外 API

```typescript
// 解析：二进制 -> 双视图 JSON
const { raw, semantic } = await parseDocx(arrayBuffer);

// 序列化：JSON -> 二进制
const buffer = await serializeDocx(semantic);
const buffer = await serializeDocx(raw);

// 仅处理 XML 片段
const rawView = parseDocxXml(xmlString);
const xml = serializeDocxXml(rawJson);
```

## 审阅信息支持

三种格式均支持解析以下审阅数据：

- **批注（Comments）**：作者、时间、内容、关联位置
- **修订记录（Track Changes）**：插入/删除/格式变更、作者、时间
- **审阅者信息**：姓名、初始、时间戳

## 技术栈

| 类别       | 选型              |
|------------|-------------------|
| 语言       | TypeScript (strict) |
| 目标       | ES2022            |
| 模块格式   | ESM only          |
| ZIP        | jszip             |
| XML        | fast-xml-parser   |
| 构建       | tsup              |
| 测试       | vitest            |
| 运行环境   | Node.js + 浏览器  |

## 测试策略

- **单元测试**：XML 解析/序列化、关系文件处理、样式计算等纯函数
- **集成测试**：完整的文件解析/序列化往返测试
- **快照测试**：用真实 Office 文件做 golden test，确保输出稳定
- 测试数据：每种格式 3-5 个样本文件（简单/中等/复杂），放在 `fixtures/` 目录

## 实现优先级

三种格式并行开发，共享 core 模块。实现顺序：

1. Core 模块（zip/xml/rels/content-type）
2. 三种格式的 parser（XML -> Raw JSON）
3. 三种格式的 semantic converter（Raw -> Semantic）
4. 三种格式的 serializer（JSON -> XML）
5. 测试与文档
