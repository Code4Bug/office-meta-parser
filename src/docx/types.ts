import type { DocumentMeta } from '../core/meta.js';
export type { DocumentMeta };

export interface DocxDocument {
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
  theme?: import('../xlsx/types.js').ThemeDefinition;
  fonts?: FontEntry[];
  settings?: DocumentSettings;
  commentExts?: CommentExtended[];
  people?: Person[];
  appMeta?: import('../core/app-meta.js').AppMeta;
  customProperties?: import('../core/custom-meta.js').CustomProperty[];
  extraParts?: Map<string, import('../core/types.js').ParsedNode>;
  extraRels?: Map<string, import('../core/types.js').Relationship[]>;
  extraContentTypes?: import('../core/types.js').ContentType[];
  extraEntries?: import('../core/types.js').ZipEntry[];
  originalRootAttrs?: Record<string, string>;
  originalStylesRootAttrs?: Record<string, string>;
  rawXmlParts?: Map<string, string>;
}

export interface Header {
  id: string;
  type?: string;
  content: Paragraph[];
}

export interface Footer {
  id: string;
  type?: string;
  content: Paragraph[];
}

export interface StyleDefinitions {
  paragraphStyles: ParagraphStyle[];
  characterStyles: CharacterStyle[];
  tableStyles: TableStyle[];
  latentStyles?: LatentStyle[];
}

export interface LatentStyle {
  name: string;
  uiPriority?: number;
  semiHidden?: boolean;
  unhideWhenUsed?: boolean;
  qFormat?: boolean;
}

export interface ParagraphStyle {
  id: string;
  name?: string;
  basedOn?: string;
  next?: string;
  isDefault?: boolean;
  uiPriority?: number;
  semiHidden?: boolean;
  unhideWhenUsed?: boolean;
  qFormat?: boolean;
  properties?: ParagraphProperties;
  runProperties?: RunProperties;
}

export interface CharacterStyle {
  id: string;
  name?: string;
  basedOn?: string;
  properties?: RunProperties;
}

export interface TableStyle {
  id: string;
  name?: string;
  basedOn?: string;
}

export interface DocxBody {
  blocks: DocxBlock[];
  sectionProperties?: SectionProperties;
}

export interface SectionProperties {
  pageWidth?: number;
  pageHeight?: number;
  marginTop?: number;
  marginRight?: number;
  marginBottom?: number;
  marginLeft?: number;
  headerMargin?: number;
  footerMargin?: number;
  gutter?: number;
  columnSpace?: number;
  columnCount?: number;
  headerReferenceId?: string;
  headerReferenceType?: string;
  footerReferenceId?: string;
  footerReferenceType?: string;
  orientation?: 'portrait' | 'landscape';
  pageNumberFormat?: string;
  pageNumberStart?: number;
  titlePage?: boolean;
  evenAndOddHeaders?: boolean;
  verticalAlign?: 'top' | 'center' | 'bottom' | 'both';
}

export type DocxBlock = Paragraph | Table | Image | Hyperlink | BookmarkStart | BookmarkEnd;

export interface Hyperlink {
  type: 'hyperlink';
  relationshipId: string;
  url?: string;
  runs: TextRun[];
  tooltip?: string;
}

export interface Paragraph {
  type: 'paragraph';
  style?: string;
  properties?: ParagraphProperties;
  runs: TextRun[];
  numbering?: NumberingProperties;
}

export interface ShadingStyle {
  fill?: string;
  color?: string;
  pattern?: string;
}

export interface ParagraphBorder {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
  between?: BorderStyle;
}

export interface TabStop {
  position: number;
  alignment: 'left' | 'right' | 'center' | 'decimal' | 'bar' | 'clear';
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface ParagraphProperties {
  alignment?: 'left' | 'center' | 'right' | 'justify';
  indent?: {
    left?: number;
    right?: number;
    start?: number;
    end?: number;
    firstLine?: number;
    hanging?: number;
    leftChars?: number;
    rightChars?: number;
    startChars?: number;
    endChars?: number;
    firstLineChars?: number;
    hangingChars?: number;
  };
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: 'auto' | 'exact' | 'atLeast';
  };
  outlineLevel?: number;
  border?: ParagraphBorder;
  shading?: ShadingStyle;
  tabs?: TabStop[];
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  formatChangeAuthor?: string;
  formatChangeDate?: string;
}

export interface NumberingProperties {
  level: number;
  numId: string;
  format?: string;
  text?: string;
}

export interface Field {
  instruction: string;
  result?: string;
}

export interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontSize?: number;
  field?: Field;
  color?: string;
  fontFamily?: string;
  properties?: RunProperties;
  highlight?: string;
  shadingColor?: string;
  shadingPattern?: string;
  caps?: boolean;
  smallCaps?: boolean;
  dstrike?: boolean;
  vanish?: boolean;
  characterSpacing?: number;
  kern?: number;
  emphasis?: string;
  characterBorder?: BorderStyle;
  verticalPosition?: number;
  revisionType?: 'insert' | 'delete';
  revisionAuthor?: string;
  revisionDate?: string;
  commentId?: string;
}

export interface RunProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  fontSizeCs?: number;
  color?: string;
  fontFamily?: string;
  fontFamilyEastAsia?: string;
  superscript?: boolean;
  subscript?: boolean;
  highlight?: string;
  shadingColor?: string;
  shadingPattern?: string;
  caps?: boolean;
  smallCaps?: boolean;
  dstrike?: boolean;
  vanish?: boolean;
  characterSpacing?: number;
  kern?: number;
  emphasis?: string;
  characterBorder?: BorderStyle;
  verticalPosition?: number;
  formatChangeAuthor?: string;
  formatChangeDate?: string;
}

export interface Table {
  type: 'table';
  rows: TableRow[];
  properties?: TableProperties;
}

export interface TableRow {
  cells: TableCell[];
  properties?: TableRowProperties;
}

export interface TableCell {
  blocks: DocxBlock[];
  properties?: TableCellProperties;
}

export interface TableProperties {
  style?: string;
  width?: number;
  widthType?: string;
  indent?: number;
  layout?: 'autofit' | 'fixed';
  cellMarginTop?: number;
  cellMarginLeft?: number;
  cellMarginBottom?: number;
  cellMarginRight?: number;
  gridColumns?: number[];
  borders?: TableBorders;
}

export interface TableRowProperties {
  height?: number;
  heightRule?: 'auto' | 'exact' | 'atLeast';
}

export interface TableCellBorders {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
}

export interface TableCellProperties {
  width?: number;
  verticalMerge?: 'restart' | 'continue';
  horizontalMerge?: 'restart' | 'continue';
  gridSpan?: number;
  verticalAlign?: 'top' | 'center' | 'bottom';
  borders?: TableCellBorders;
  noWrap?: boolean;
}

export interface TableBorders {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
  insideHorizontal?: BorderStyle;
  insideVertical?: BorderStyle;
}

export interface BorderStyle {
  style: string;
  size?: number;
  color?: string;
}

export interface Image {
  type: 'image';
  relationshipId: string;
  width?: number;
  height?: number;
  alt?: string;
  isFloating?: boolean;
  wrapType?: 'square' | 'tight' | 'through' | 'topBottom' | 'none' | 'behind' | 'front';
  posX?: number;
  posY?: number;
}

export interface Comment {
  id: string;
  author: string;
  date: string;
  content: Paragraph[];
  initials?: string;
}

export interface Revision {
  type: 'insert' | 'delete' | 'formatChange';
  author: string;
  date: string;
}

export interface NumberingLevel {
  level: number;
  numFmt: string;
  lvlText: string;
  start: number;
  indent?: { left?: number; hanging?: number };
  alignment?: string;
  runProperties?: RunProperties;
}

export interface AbstractNum {
  id: string;
  levels: NumberingLevel[];
}

export interface Num {
  id: string;
  abstractNumId: string;
}

export interface NumberingDefinitions {
  abstractNums: AbstractNum[];
  nums: Num[];
}

export interface Footnote {
  id: string;
  type?: string;
  content: Paragraph[];
}

export interface BookmarkStart {
  type: 'bookmarkStart';
  id: string;
  name: string;
}

export interface BookmarkEnd {
  type: 'bookmarkEnd';
  id: string;
}

export interface FontEntry {
  name: string;
  altName?: string;
  charset?: string;
  family?: string;
  pitch?: string;
}

export interface DocumentSettings {
  defaultTabStop?: number;
  zoom?: number;
  compatibilityMode?: number;
  evenAndOddHeaders?: boolean;
  documentProtection?: boolean;
  characterSpacingControl?: string;
}

export interface CommentExtended {
  paraId: string;
  done?: boolean;
}

export interface Person {
  author: string;
  userId?: string;
  providerId?: string;
}
