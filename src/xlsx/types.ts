import type { DocumentMeta } from '../core/meta.js';
import type { AppMeta } from '../core/app-meta.js';
import type { CustomProperty } from '../core/custom-meta.js';

export interface SharedStringEntry {
  text: string;
  richText?: RichTextRun[];
}

export interface XlsxWorkbook {
  meta: DocumentMeta;
  sheets: Sheet[];
  styles: CellStyleDefinitions;
  sharedStrings: SharedStringEntry[];
  definedNames?: DefinedName[];
  theme?: ThemeDefinition;
  appMeta?: AppMeta;
  customProperties?: CustomProperty[];
  extraParts?: Map<string, import('../core/types.js').ParsedNode>;
  extraRels?: Map<string, import('../core/types.js').Relationship[]>;
  extraContentTypes?: import('../core/types.js').ContentType[];
  extraEntries?: import('../core/types.js').ZipEntry[];
  rawXmlParts?: Map<string, string>;
}

export interface Sheet {
  name: string;
  cells: Cell[][];
  mergedCells: MergedCell[];
  columnWidths: number[];
  rowHeights: number[];
  hyperlinks: Hyperlink[];
  autoFilter?: AutoFilter;
  dataValidations?: DataValidation[];
  conditionalFormats?: ConditionalFormat[];
  printArea?: PrintArea;
  images?: SheetImage[];
  state?: 'visible' | 'hidden' | 'veryHidden';
  tabColor?: string;
  frozenPanes?: FrozenPanes;
  zoomScale?: number;
  activeCell?: string;
  selections?: SheetSelection[];
  rowGroups?: OutlineGroup[];
  colGroups?: OutlineGroup[];
  defaultRowHeight?: number;
  defaultColWidth?: number;
  pageSetup?: PageSetup;
  headerFooter?: SheetHeaderFooter;
  comments?: SheetComment[];
  tables?: ExcelTable[];
  protection?: SheetProtection;
  printTitles?: string;
}

export interface Hyperlink {
  ref: string;
  url: string;
  tooltip?: string;
}

export interface Cell {
  value: string | number | boolean | null;
  formula?: string;
  formulaType?: 'normal' | 'array' | 'shared';
  formulaRef?: string;
  sharedFormulaIndex?: number;
  style?: string;
  type: 'string' | 'number' | 'boolean' | 'formula' | 'sharedString' | 'date' | 'error';
  richText?: RichTextRun[];
}

export interface RichTextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  size?: number;
  font?: string;
}

export interface MergedCell {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface CellStyleDefinitions {
  cellStyles: CellStyle[];
  fonts: FontDefinition[];
  fills: FillDefinition[];
  borders: BorderDefinition[];
  numberFormats: NumberFormatDefinition[];
}

export interface CellStyle {
  id: string;
  font?: Partial<FontDefinition>;
  fill?: Partial<FillDefinition>;
  border?: Partial<BorderDefinition>;
  alignment?: Alignment;
  numberFormat?: string;
  protection?: CellProtection;
}

export interface FontDefinition {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  vertAlign?: 'superscript' | 'subscript';
  scheme?: 'major' | 'minor';
}

export interface FillDefinition {
  patternType?: string;
  fgColor?: string;
  bgColor?: string;
  gradientFill?: GradientFill;
}

export interface GradientFill {
  type?: 'linear' | 'path';
  degree?: number;
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  stops: GradientStop[];
}

export interface GradientStop {
  position: number;
  color: string;
}

export interface BorderDefinition {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
  diagonal?: BorderStyle;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export interface BorderStyle {
  style?: string;
  color?: string;
}

export interface Alignment {
  horizontal?: 'left' | 'center' | 'right' | 'justify';
  vertical?: 'top' | 'center' | 'bottom';
  wrapText?: boolean;
  indent?: number;
  textRotation?: number;
  shrinkToFit?: boolean;
  readingOrder?: 0 | 1 | 2;
  justifyLastLine?: boolean;
}

export interface CellProtection {
  locked?: boolean;
  hidden?: boolean;
}

export interface SheetProtection {
  enabled?: boolean;
  password?: string;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  insertRows?: boolean;
  deleteRows?: boolean;
  formatCells?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
}

export interface NumberFormatDefinition {
  id: string;
  formatCode: string;
}

export interface AutoFilter {
  ref: string;
  columns: AutoFilterColumn[];
}

export interface AutoFilterColumn {
  colId: number;
  filters?: string[];
  operator?: string;
  customFilter?: { operator: string; value: string };
}

export interface DataValidation {
  type: string;
  operator?: string;
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  sqref: string;
  formula1?: string;
  formula2?: string;
}

export interface ConditionalFormat {
  sqref: string;
  rules: ConditionalRule[];
}

export interface ConditionalRule {
  type: string;
  priority: number;
  formula?: string[];
  colorScale?: { colors: string[] };
  dataBar?: { color: string };
  iconSet?: { iconSet: string };
  dxfId?: number;
}

export interface PrintArea {
  sheet?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  pageMargins?: { top: number; right: number; bottom: number; left: number; header: number; footer: number };
}

export interface SheetImage {
  relationshipId: string;
  name?: string;
  description?: string;
  position: {
    from: { col: number; row: number; colOff?: number; rowOff?: number };
    to: { col: number; row: number; colOff?: number; rowOff?: number };
  };
  size?: { width: number; height: number };
}

export interface FrozenPanes {
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
}

export interface SheetSelection {
  pane?: string;
  activeCell?: string;
  sqref?: string;
}

export interface OutlineGroup {
  level: number;
  collapsed?: boolean;
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  paperSize?: number;
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  firstPageNumber?: number;
  useFirstPageNumber?: boolean;
  horizontalDpi?: number;
  verticalDpi?: number;
}

export interface SheetHeaderFooter {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  differentFirst?: boolean;
  differentOddEven?: boolean;
}

export interface SheetComment {
  ref: string;
  authorId: number;
  text: string;
  richText?: RichTextRun[];
}

export interface ExcelTable {
  id: number;
  name: string;
  displayName: string;
  ref: string;
  headerRowCount?: number;
  totalsRowCount?: number;
  columns: TableColumn[];
}

export interface TableColumn {
  id: number;
  name: string;
  totalsRowLabel?: string;
  totalsRowFunction?: string;
}

export interface DefinedName {
  name: string;
  formula: string;
  localSheetId?: number;
  hidden?: boolean;
}

export interface ThemeDefinition {
  colorScheme?: Record<string, string>;
  majorFont?: string;
  minorFont?: string;
}
