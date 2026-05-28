import type { DocumentMeta } from '../core/meta.js';
import type { Paragraph } from '../docx/types.js';
import type { AppMeta } from '../core/app-meta.js';
import type { CustomProperty } from '../core/custom-meta.js';

export interface PptxPresentation {
  meta: DocumentMeta;
  slides: Slide[];
  masters: SlideMaster[];
  layouts: SlideLayout[];
  theme?: Theme;
  slideSize?: { width: number; height: number };
  notesSize?: { width: number; height: number };
  appMeta?: AppMeta;
  customProperties?: CustomProperty[];
  rawXmlParts?: Map<string, string>;
}

export interface Slide {
  elements: SlideElement[];
  transition?: Transition;
  notes?: string;
  layout?: string;
  animations?: Animation[];
  background?: FillStyle;
  showMasterSp?: boolean;
  showMasterPhAnim?: boolean;
  clrMap?: Record<string, string>;
  comments?: SlideComment[];
}

export interface SlideComment {
  id: string;
  authorId: number;
  authorName: string;
  text: string;
  date?: string;
  position?: { x: number; y: number };
  replies?: SlideComment[];
}

export type SlideElement = TextShape | ImageShape | GroupShape | TableShape | MediaShape;

export interface TextShape {
  type: 'text';
  content: string;
  position: Position;
  style?: ShapeStyle;
  paragraphs: Paragraph[];
  rotation?: number;
  placeholder?: {
    type: string;
    index?: number;
  };
  hyperlink?: Hyperlink;
  presetGeom?: string;
  bodyProperties?: BodyProperties;
  listStyle?: ListStyle;
}

export interface BodyProperties {
  anchor?: 't' | 'ctr' | 'b';
  wrap?: 'square' | 'none';
  leftInset?: number;
  topInset?: number;
  rightInset?: number;
  bottomInset?: number;
  vertical?: 'horz' | 'vert' | 'vert270' | 'wordArtVert';
  autoFit?: 'none' | 'normal' | 'shape';
}

export interface ListStyle {
  defaultParagraphProperties?: ListLevelProperties[];
}

export interface ListLevelProperties {
  level: number;
  alignment?: string;
  indent?: number;
  marL?: number;
  fontScale?: number;
}

export interface ImageShape {
  type: 'image';
  relationshipId: string;
  position: Position;
  style?: ShapeStyle;
  alt?: string;
  rotation?: number;
  hyperlink?: Hyperlink;
}

export interface Hyperlink {
  url: string;
  tooltip?: string;
}

export interface GroupShape {
  type: 'group';
  position: Position;
  children: SlideElement[];
  childOffset?: { x: number; y: number };
  childExtent?: { width: number; height: number };
}

export interface PptxTableCellStyle {
  fill?: FillStyle;
  borders?: {
    top?: BorderStyle;
    bottom?: BorderStyle;
    left?: BorderStyle;
    right?: BorderStyle;
  };
}

export interface TableShape {
  type: 'table';
  position: Position;
  rows: PptxTableRow[];
  style?: ShapeStyle;
  tableStyleId?: string;
}

export interface PptxTableRow {
  cells: PptxTableCell[];
  height?: number;
}

export interface PptxTableCell {
  content: Paragraph[];
  width?: number;
  colspan?: number;
  rowspan?: number;
  style?: PptxTableCellStyle;
}

export interface MediaShape {
  type: 'media';
  mediaType: 'video' | 'audio';
  relationshipId: string;
  position: Position;
}

export interface Position {
  x: number;
  y: number;
  width: number;
  height: number;
  flipH?: boolean;
  flipV?: boolean;
  rotation?: number;
}

export interface ShapeStyle {
  fill?: FillStyle;
  border?: BorderStyle;
  shadow?: ShadowStyle;
  opacity?: number;
  glow?: GlowStyle;
  softEdge?: SoftEdgeStyle;
  reflection?: ReflectionStyle;
}

export interface GlowStyle {
  color?: string;
  radius?: number;
}

export interface SoftEdgeStyle {
  radius?: number;
}

export interface FillStyle {
  color?: string;
  type?: 'solid' | 'gradient' | 'pattern' | 'none' | 'blip' | 'group';
  gradientFill?: GradientFill;
  patternFill?: PatternFill;
  blipRelationshipId?: string;
}

export interface GradientFill {
  type?: 'linear' | 'path';
  angle?: number;
  stops: GradientStop[];
}

export interface GradientStop {
  position: number;
  color: string;
}

export interface PatternFill {
  preset?: string;
  fgColor?: string;
  bgColor?: string;
}

export interface BorderStyle {
  color?: string;
  width?: number;
  style?: 'solid' | 'dashed' | 'dotted';
  dashType?: string;
  headEnd?: LineEnd;
  tailEnd?: LineEnd;
  compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
  cap?: 'flat' | 'round' | 'sq';
}

export interface LineEnd {
  type?: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
  width?: 'sm' | 'med' | 'lg';
  length?: 'sm' | 'med' | 'lg';
}

export interface ReflectionStyle {
  blur?: number;
  distance?: number;
  startOpacity?: number;
  endOpacity?: number;
  direction?: number;
  scaleY?: number;
}

export interface ShadowStyle {
  type?: 'outer' | 'inner';
  color?: string;
  blur?: number;
  offsetX?: number;
  offsetY?: number;
}

export interface SlideMaster {
  id: string;
  layouts: SlideLayout[];
  background?: FillStyle;
  txStyles?: MasterTxStyles;
}

export interface MasterTxStyles {
  titleStyle?: TextStyle;
  bodyStyle?: TextStyle;
  otherStyle?: TextStyle;
}

export interface TextStyle {
  levels?: TextLevelStyle[];
}

export interface TextLevelStyle {
  level: number;
  alignment?: string;
  marL?: number;
  indent?: number;
  fontScale?: number;
  spcBef?: number;
  spcAft?: number;
  lnSpc?: number;
}

export interface SlideLayout {
  id: string;
  name?: string;
  type?: string;
  placeholders: Placeholder[];
}

export interface Placeholder {
  type: string;
  position: Position;
  index?: number;
}

export interface Transition {
  type: string;
  duration?: number;
  direction?: string;
  advClick?: boolean;
  advTime?: number;
  sound?: string;
}

export interface Theme {
  colorScheme: ColorScheme;
  fontScheme: FontScheme;
  formatScheme?: FormatScheme;
}

export interface FormatScheme {
  fillStyles?: FillStyle[];
  lineStyles?: BorderStyle[];
  effectStyles?: EffectStyle[];
  bgFillStyles?: FillStyle[];
}

export interface EffectStyle {
  shadow?: ShadowStyle;
  glow?: GlowStyle;
  softEdge?: SoftEdgeStyle;
}

export interface ColorScheme {
  name?: string;
  colors: Record<string, string>;
}

export interface FontScheme {
  name?: string;
  majorFont?: string;
  minorFont?: string;
}

export interface Animation {
  trigger: 'onClick' | 'withPrevious' | 'afterPrevious' | 'onLoad';
  type: string;
  shapeId?: string;
  duration?: number;
  delay?: number;
  direction?: string;
  children?: Animation[];
}
