# office-meta-parser P0+P1 功能全量实现规格

> 日期：2026-05-25  
> 范围：DOCX P0 (19项) + P1 (13项)、XLSX P0 (22项) + P1 (5项)、PPTX P0 (30项) + P1 (8项)，共 97 项  
> 实现顺序：DOCX P0 → DOCX P1 → XLSX P0 → XLSX P1 → PPTX P0 → PPTX P1  
> 方案：方案 A（按格式 + 优先级分批，每批一次提交）

---

## 一、总体架构原则

每项功能都遵循统一三步模式：
1. **类型定义**：在 `types.ts` 新增/扩展接口字段
2. **解析器**：在 `parsers/*.ts` 从 XML 节点读取属性
3. **序列化器**：在 `serializers/*.ts` 将类型字段写回 XML 节点

文件行数控制在 300 行以内；若扩展后超限，拆分为新子文件。

---

## 二、DOCX P0（19项）

### 2.1 RunProperties 扩展（10 个新字段）

在 `src/docx/types.ts` 的 `RunProperties` 接口新增：

```ts
highlight?: string;           // w:highlight@w:val（颜色名）
shadingColor?: string;        // w:shd@w:fill（run 底纹）
shadingPattern?: string;      // w:shd@w:val
caps?: boolean;               // w:caps
smallCaps?: boolean;          // w:smallCaps
dstrike?: boolean;            // w:dstrike（双删除线）
vanish?: boolean;             // w:vanish（隐藏）
characterSpacing?: number;    // w:spacing@w:val（emu）
kern?: number;                // w:kern@w:val
```

解析：扩展 `parsers/paragraph.ts` 中 `parseRunProperties`，新增 switch-case 分支。  
序列化：扩展 `serializers/paragraph.ts` 中 `serializeRunProperties`，按相同模式生成节点。

### 2.2 ParagraphProperties 扩展（7 个新字段）

在 `ParagraphProperties` 接口新增：

```ts
outlineLevel?: number;        // w:outlineLvl@w:val
border?: ParagraphBorder;     // w:pBdr
shading?: ShadingStyle;       // w:shd（段落底纹）
tabs?: TabStop[];             // w:tabs/w:tab
keepNext?: boolean;           // w:keepNext
keepLines?: boolean;          // w:keepLines
pageBreakBefore?: boolean;    // w:pageBreakBefore
```

新类型：
```ts
interface ParagraphBorder {
  top?: BorderStyle; bottom?: BorderStyle;
  left?: BorderStyle; right?: BorderStyle;
  between?: BorderStyle;
}
interface ShadingStyle {
  fill?: string; color?: string; pattern?: string;
}
interface TabStop {
  position: number;       // w:tab@w:pos（twips）
  alignment: 'left' | 'right' | 'center' | 'decimal' | 'bar' | 'clear';
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}
```

### 2.3 SectionProperties 扩展（4 个新字段）

```ts
orientation?: 'portrait' | 'landscape';   // w:pgSz@w:orient
pageNumberFormat?: string;                 // w:pgNumType@w:fmt
pageNumberStart?: number;                  // w:pgNumType@w:start
titlePage?: boolean;                       // w:titlePg
evenAndOddHeaders?: boolean;               // w:evenAndOddHeaders（body level）
verticalAlign?: 'top' | 'center' | 'bottom' | 'both'; // w:vAlign@w:val
```

### 2.4 完整列表定义（numbering.xml）

**新文件** `src/docx/parsers/numbering.ts`：

```ts
interface NumberingDefinitions {
  abstractNums: AbstractNum[];
  nums: Num[];
}
interface AbstractNum {
  id: string;
  levels: NumberingLevel[];
}
interface NumberingLevel {
  level: number;
  numFmt: string;           // w:numFmt@w:val
  lvlText: string;          // w:lvlText@w:val
  start: number;
  indent?: { left?: number; hanging?: number };
  alignment?: string;
  runProperties?: RunProperties;
}
interface Num {
  id: string;
  abstractNumId: string;
}
```

`DocxDocument` 新增 `numbering?: NumberingDefinitions`。  
在 `parser.ts` 中读取 `word/numbering.xml` 并解析。  
**新文件** `src/docx/serializers/numbering.ts`：将 `NumberingDefinitions` 序列化为 `word/numbering.xml`。

### 2.5 脚注（footnotes.xml）

**新文件** `src/docx/parsers/footnote.ts`：

```ts
interface Footnote {
  id: string;
  content: Paragraph[];
}
```

`DocxDocument` 新增 `footnotes?: Footnote[]`。  
在 `parser.ts` 中读取 `word/footnotes.xml`。  
TextRun 新增 `footnoteRef?: string` 字段（引用脚注 id）。  
**新文件** `src/docx/serializers/footnote.ts`。

### 2.6 书签（bookmark）

`DocxBlock` 扩展新类型 `BookmarkStart | BookmarkEnd`：

```ts
interface BookmarkStart { type: 'bookmarkStart'; id: string; name: string; }
interface BookmarkEnd   { type: 'bookmarkEnd'; id: string; }
```

在 `parsers/body.ts` 中识别 `w:bookmarkStart / w:bookmarkEnd`。  
在 `serializers/document.ts` 中序列化书签节点。

### 2.7 完整跨 run 域结构（fldChar）

当前实现已支持基本字段。扩展到支持：
- 嵌套字段（field 内有 field）
- 多 run 跨越的 instrText 拼接

在 `parsers/paragraph.ts` 的字段状态机中，正确累积跨多个 run 的 instrText。

---

## 三、DOCX P1（13项）

### 3.1 浮动图片（wp:anchor）

`Image` 类型新增 `anchor` 标志：

```ts
interface Image {
  // ...existing
  floating?: boolean;          // wp:anchor（vs wp:inline）
  wrapType?: 'square' | 'tight' | 'through' | 'topBottom' | 'none';
  posH?: { relativeFrom: string; offset?: number; align?: string };
  posV?: { relativeFrom: string; offset?: number; align?: string };
}
```

在 `parsers/image.ts` 和 `serializers/image.ts` 中扩展。

### 3.2 图片旋转/翻转/裁剪

在 `Image` 接口新增：
```ts
rotation?: number;
flipH?: boolean;
flipV?: boolean;
cropLeft?: number; cropRight?: number; cropTop?: number; cropBottom?: number;
```

从 `a:xfrm@rot`、`flipH/V`、`a:srcRect` 解析。

### 3.3 单元格底纹、独立边框、noWrap、tblLook

在 `TableCellProperties` 新增：
```ts
shading?: ShadingStyle;
borders?: TableBorders;
noWrap?: boolean;
```

在 `TableProperties` 新增：
```ts
look?: { firstRow?: boolean; firstCol?: boolean; lastRow?: boolean; lastCol?: boolean; bandRow?: boolean; bandCol?: boolean };
```

### 3.4 RunProperties P1 字段

在 `RunProperties` 新增：
```ts
emphasisMark?: string;    // w:em@w:val（中日韩着重号）
lang?: string;            // w:lang@w:val
border?: BorderStyle;     // w:bdr
position?: number;        // w:position@w:val（半磅垂直偏移）
```

---

## 四、XLSX P0（22项）

### 4.1 SheetView 扩展（4项）

在 `src/xlsx/types.ts` 的 Sheet 接口新增：

```ts
interface SheetView {
  zoomScale?: number;
  topLeftCell?: string;
  activeCell?: string;
  pane?: FrozenPane;
  selection?: { activeCell?: string; sqref?: string };
}
interface FrozenPane {
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  state: 'frozen' | 'frozenSplit' | 'split';
  activePane?: string;
}
```

解析 `sheetView/pane`、`sheetView/selection`，序列化回相同结构。

### 4.2 工作表/工作簿级元数据

```ts
// Sheet
hiddenState?: 'visible' | 'hidden' | 'veryHidden';  // sheet@state
tabColor?: string;                                    // sheetPr/tabColor@rgb

// Sheet
groupRows?: RowGroup[];    // row@outlineLevel + outlinePr
groupCols?: ColGroup[];

// SheetFormatPr
defaultRowHeight?: number;
defaultColWidth?: number;
```

### 4.3 完整 pageSetup（1项）

扩展现有 `PrintOptions` 类型加入：
```ts
orientation?: 'portrait' | 'landscape';
paperSize?: number;
scale?: number;
```

### 4.4 页眉页脚（1项）

```ts
interface HeaderFooter {
  oddHeader?: string;   // 支持 &P &N &D &F &A 占位符
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
}
```

Sheet 新增 `headerFooter?: HeaderFooter`。

### 4.5 命名区域（definedNames）

```ts
interface DefinedName {
  name: string;
  value: string;    // 如 "Sheet1!$A$1:$B$10"
  localSheetId?: number;
  comment?: string;
}
```

`XlsxWorkbook` 新增 `definedNames?: DefinedName[]`。

### 4.6 Excel Tables（tableParts）

```ts
interface ExcelTable {
  id: string;
  name: string;
  ref: string;        // "A1:D10"
  displayName: string;
  columns: TableColumn[];
  autoFilter?: boolean;
}
interface TableColumn { id: string; name: string; }
```

Sheet 新增 `tables?: ExcelTable[]`。  
在 `parser.ts` 中读取 `xl/tables/table*.xml`。

### 4.7 批注（2项）

**新文件** `src/xlsx/parsers/comments.ts`：
```ts
interface CellComment {
  ref: string;        // 如 "A1"
  author: string;
  text: string;
  isThreaded?: boolean;
}
```

Sheet 新增 `comments?: CellComment[]`。  
解析 `xl/comments*.xml` 中的 `<comment>` 元素。

### 4.8 富文本字符串（sharedStrings）

`SharedString` 从 `string` 扩展为：
```ts
type SharedString = string | RichTextString;
interface RichTextString {
  runs: RichTextRun[];
}
interface RichTextRun {
  text: string;
  bold?: boolean; italic?: boolean; underline?: boolean;
  fontSize?: number; color?: string; fontFamily?: string;
  vertAlign?: 'superscript' | 'subscript';
}
```

### 4.9 样式完整扩展（9项）

扩展 `src/xlsx/types.ts` 中的样式类型：

```ts
// Alignment 新增
indent?: number;
textRotation?: number;       // -90 ~ 90，255 = 竖排
shrinkToFit?: boolean;

// Fill 新增
gradientFill?: GradientFill;
interface GradientFill {
  degree?: number;
  stops: Array<{ position: number; color: string }>;
}

// Border 新增
diagonal?: BorderSide;
diagonalUp?: boolean;
diagonalDown?: boolean;

// Font 新增
vertAlign?: 'superscript' | 'subscript' | 'baseline';
scheme?: 'major' | 'minor';

// Font.vertAlign 单独从 cell 上标下标拆出
```

### 4.10 Theme1.xml（颜色主题）

```ts
interface XlsxTheme {
  colors: Record<string, string>;   // dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink
}
```

`XlsxWorkbook` 新增 `theme?: XlsxTheme`。  
解析 `xl/theme/theme1.xml` 中的 `<a:clrScheme>`。

---

## 五、XLSX P1（5项）

```ts
// 打印标题 (definedName _xlnm.Print_Titles)
// → 已通过 definedNames 覆盖，无需额外类型

// 工作表保护
interface SheetProtection {
  sheet?: boolean; password?: string;
  selectLockedCells?: boolean; selectUnlockedCells?: boolean;
}
// Sheet 新增 protection?: SheetProtection

// 单元格保护
// CellStyle 新增 locked?: boolean; hidden?: boolean;

// readingOrder / justifyLastLine
// Alignment 新增 readingOrder?: 0 | 1 | 2; justifyLastLine?: boolean;
```

---

## 六、PPTX P0（30项）

### 6.1 TextBodyProperties 新类型

```ts
interface TextBodyProperties {
  anchor?: 't' | 'ctr' | 'b' | 'just' | 'dist';  // a:bodyPr@anchor
  lIns?: number; tIns?: number; rIns?: number; bIns?: number;
  wrap?: 'none' | 'square';
  autofit?: 'none' | 'spAutoFit' | 'normAutoFit';
  vert?: 'horz' | 'vert' | 'vert270' | 'wordArtVert';
}
```

`TextShape` 新增 `bodyProperties?: TextBodyProperties`。

### 6.2 PptxParagraphProperties 新类型

当前 PPTX TextShape 复用 DOCX Paragraph 类型。新增独立的段落属性：

```ts
interface PptxParagraphProperties {
  alignment?: 'l' | 'ctr' | 'r' | 'just' | 'dist';
  marL?: number; indent?: number;
  level?: number;
  spaceBefore?: number; spaceAfter?: number;
  lineSpacing?: number; lineSpacingType?: 'spcPts' | 'spcPct';
  bullet?: BulletStyle;
  defaultRunProps?: PptxRunProperties;
}
interface BulletStyle {
  type: 'none' | 'char' | 'autoNum' | 'pic';
  char?: string;
  autoNumType?: string;
  font?: string;
  sizePct?: number;
  color?: string;
}
```

### 6.3 PPTX 段落/run 类型与 DOCX 的解耦策略

**设计决策**：在 `src/pptx/types.ts` 新增独立的 `PptxParagraph` 和 `PptxRunProperties`，不再复用 DOCX 的 `Paragraph` 类型。`TextShape.paragraphs` 类型由 `Paragraph[]`（DOCX）改为 `PptxParagraph[]`。

原因：a:pPr 和 w:pPr 的属性集差异显著（对齐值、项目符号、行距等），共用 DOCX 类型会引入不匹配的字段，且后续扩展受限。

```ts
interface PptxParagraph {
  type: 'paragraph';
  properties?: PptxParagraphProperties;
  runs: PptxTextRun[];
}
interface PptxTextRun {
  text: string;
  properties?: PptxRunProperties;
}
interface PptxRunProperties {
  bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean;
  fontSize?: number; color?: string; fontFamily?: string;
  superscript?: boolean; subscript?: boolean;
  highlight?: string; lang?: string;
}
```

`TextShape` 修改为：
```ts
interface TextShape {
  // ...
  paragraphs: PptxParagraph[];   // 替换原来的 Paragraph[]（DOCX）
}
```

同时需要更新 PPTX 的 parsers/serializers 中所有引用 DOCX `Paragraph` 的地方。

### 6.4 Position 扩展（flipH/V、GroupShape 子坐标）

```ts
interface Position {
  x: number; y: number; width: number; height: number;
  rotation?: number;
  flipH?: boolean;   // a:xfrm@flipH
  flipV?: boolean;   // a:xfrm@flipV
}
// GroupShape 新增
interface GroupShape {
  type: 'group';
  position: Position;
  childOffset?: { x: number; y: number };   // a:chOff
  childExtent?: { cx: number; cy: number }; // a:chExt
  rotation?: number;                         // grpSpPr/a:xfrm@rot
  children: SlideElement[];
}
```

### 6.5 完整 FillStyle

```ts
interface FillStyle {
  type?: 'solid' | 'gradient' | 'pattern' | 'none' | 'group' | 'blip';
  color?: string;
  gradient?: GradientFill;
  pattern?: PatternFill;
  blipRelId?: string;       // 形状图片填充
}
interface GradientFill {
  angle?: number;
  stops: Array<{ pos: number; color: string }>;
  path?: 'circle' | 'rect' | 'shape';
}
interface PatternFill {
  preset: string;
  fgColor?: string;
  bgColor?: string;
}
```

### 6.6 完整 BorderStyle（线条属性）

```ts
interface BorderStyle {
  color?: string; width?: number;
  style?: 'solid' | 'dashed' | 'dotted';
  cap?: 'flat' | 'round' | 'sq';
  compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
  headEnd?: ArrowEnd;
  tailEnd?: ArrowEnd;
}
interface ArrowEnd {
  type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
  width?: 'sm' | 'med' | 'lg';
  length?: 'sm' | 'med' | 'lg';
}
```

### 6.7 ShapeStyle 效果扩展

```ts
interface ShapeStyle {
  // ...existing
  glow?: { radius: number; color: string };
  softEdge?: { radius: number };
  preset?: string;   // a:prstGeom@prst
}
```

### 6.8 Theme 完整结构（fmtScheme）

```ts
interface Theme {
  colorScheme?: Record<string, string>;
  fontScheme?: { major?: string; minor?: string };
  fmtScheme?: {
    fillStyleLst?: FillStyle[];
    lnStyleLst?: BorderStyle[];
    effectStyleLst?: EffectStyle[];
    bgFillStyleLst?: FillStyle[];
  };
}
```

### 6.9 SlideMaster txStyles

`SlideMaster` 新增：
```ts
txStyles?: {
  titleStyle?: PptxParagraphProperties;
  bodyStyle?: PptxParagraphProperties;
  otherStyle?: PptxParagraphProperties;
}
```

### 6.10 颜色映射（clrMap / clrMapOvr）

```ts
interface ColorMap {
  bg1?: string; tx1?: string; bg2?: string; tx2?: string;
  accent1?: string; accent2?: string; accent3?: string;
  accent4?: string; accent5?: string; accent6?: string;
  hlink?: string; folHlink?: string;
}
```

`SlideMaster`、`SlideLayout`、`Slide` 各新增 `colorMap?: ColorMap`。

### 6.11 幻灯片尺寸

```ts
PptxPresentation 新增：
  slideSize?: { width: number; height: number };
  notesSize?: { width: number; height: number };
```

解析 `p:sldSz`，序列化时输出（替换当前硬编码值）。

### 6.12 Slide 独立背景

```ts
Slide 新增：
  background?: { fill?: FillStyle; color?: string };
```

### 6.13 母版形状显示标志

```ts
Slide 新增：
  showMasterSp?: boolean;
  showMasterPhAnim?: boolean;
```

---

## 七、PPTX P1（8项）

### 7.1 表格完整属性

```ts
interface PptxTableCell {
  // ...existing
  style?: { fill?: FillStyle; borders?: TableBorders };
  // borders: top/bottom/left/right/insideH/insideV
}
interface PptxTableStyle {
  id?: string;
  firstRow?: boolean; firstCol?: boolean;
  lastRow?: boolean; lastCol?: boolean;
  bandRow?: boolean; bandCol?: boolean;
}
TableShape 新增 tableStyle?: PptxTableStyle;
```

### 7.2 视觉效果（反射、内阴影）

```ts
ShapeStyle 新增：
  reflection?: { blurRadius?: number; dist?: number; direction?: number; fadeDir?: number; alpha?: number };
  innerShadow?: { blurRadius?: number; dist?: number; direction?: number; color?: string };
```

### 7.3 完整 dash 类型

`BorderStyle.style` 改为完整枚举：
```ts
style?: 'solid' | 'dash' | 'dot' | 'dashDot' | 'dashDotDot' | 'sysDash' | 'sysDot' | 'sysDashDot';
```

### 7.4 切换详细属性

```ts
interface Transition {
  // ...existing
  advClick?: boolean;
  advTm?: number;      // ms
  sound?: string;      // 音效文件 relId
}
```

### 7.5 嵌入音视频

`SlideElement` 新增 `MediaShape`：
```ts
interface MediaShape {
  type: 'video' | 'audio';
  relationshipId: string;
  position: Position;
  poster?: string;
}
```

---

## 八、文件影响范围汇总

| 文件 | 变更类型 |
|------|----------|
| `src/docx/types.ts` | 扩展 RunProperties、ParagraphProperties、SectionProperties；新增 Footnote、Bookmark、TabStop 等 |
| `src/docx/parsers/paragraph.ts` | 扩展 parseRunProperties、parseParagraphProperties |
| `src/docx/serializers/paragraph.ts` | 扩展 serializeRunProperties、serializeParagraphProperties |
| `src/docx/parsers/numbering.ts` | **新文件** |
| `src/docx/serializers/numbering.ts` | **新文件** |
| `src/docx/parsers/footnote.ts` | **新文件** |
| `src/docx/serializers/footnote.ts` | **新文件** |
| `src/docx/parsers/body.ts` | 扩展 bookmark 识别 |
| `src/docx/parser.ts` | 读取 numbering.xml、footnotes.xml |
| `src/docx/serializer.ts` | 序列化 numbering、footnotes |
| `src/docx/parsers/image.ts` | 扩展 anchor、rotation、flip、crop |
| `src/docx/parsers/table.ts` | 扩展 cell 底纹、独立边框、noWrap、tblLook |
| `src/xlsx/types.ts` | 扩展 SheetView/FrozenPane、样式全套、新增 DefinedName、ExcelTable、CellComment、XlsxTheme 等 |
| `src/xlsx/parsers/sheet.ts` | 扩展视图、冻结、分组、默认尺寸 |
| `src/xlsx/parsers/styles.ts` | 扩展 Alignment/Fill/Border/Font 完整字段 |
| `src/xlsx/parsers/comments.ts` | **新文件** |
| `src/xlsx/parsers/strings.ts` | 扩展富文本 |
| `src/xlsx/parser.ts` | 读取 theme1.xml、table*.xml、comments*.xml |
| `src/xlsx/serializers/worksheet.ts` | 序列化视图、冻结、分组等 |
| `src/xlsx/serializers/styles.ts` | 序列化完整样式 |
| `src/pptx/types.ts` | 新增 TextBodyProperties、PptxParagraphProperties、PptxRunProperties、完整 FillStyle/BorderStyle；扩展 Position、Theme、ColorMap 等 |
| `src/pptx/parsers/shape.ts` | 解析 bodyPr、pPr、完整 fill/border |
| `src/pptx/parsers/style.ts` | 解析 fmtScheme、完整效果 |
| `src/pptx/parsers/master.ts` | 解析 txStyles、colorMap |
| `src/pptx/parsers/theme.ts` | 解析 fmtScheme |
| `src/pptx/serializers/slide.ts` | 序列化 bodyPr、pPr、颜色映射等 |
| `src/pptx/serializers/master.ts` | 序列化 txStyles |
| `src/pptx/serializers/theme.ts` | 序列化 fmtScheme |
| `src/pptx/serializers/presentation.ts` | 序列化 sldSz/notesSz |

---

## 九、测试策略

- 每个新字段至少一个 parse → serialize → re-parse round-trip 单元测试
- numbering、footnote、table 等新文件有独立测试文件
- 行数超 300 行的文件拆分后更新测试引用
