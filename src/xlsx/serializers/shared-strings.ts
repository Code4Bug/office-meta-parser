import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { SharedStringEntry, RichTextRun } from '../types.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export function serializeSharedStrings(strings: SharedStringEntry[]): string {
  const children = strings.map(s => {
    if (s.richText && s.richText.length > 0) {
      return {
        tag: 'si',
        attrs: {},
        children: s.richText.map(run => serializeRichTextRun(run)),
      };
    }
    return {
      tag: 'si',
      attrs: {},
      children: [{ tag: 't', attrs: {}, children: [s.text] }],
    };
  });

  const root: ParsedNode = {
    tag: 'sst',
    attrs: { xmlns: MAIN_NS, count: String(strings.length), uniqueCount: String(strings.length) },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeRichTextRun(run: RichTextRun): ParsedNode {
  const rPrChildren: ParsedNode[] = [];
  if (run.bold) rPrChildren.push({ tag: 'b', attrs: {}, children: [] });
  if (run.italic) rPrChildren.push({ tag: 'i', attrs: {}, children: [] });
  if (run.underline) rPrChildren.push({ tag: 'u', attrs: {}, children: [] });
  if (run.strike) rPrChildren.push({ tag: 'strike', attrs: {}, children: [] });
  if (run.size) rPrChildren.push({ tag: 'sz', attrs: { val: String(run.size) }, children: [] });
  if (run.font) rPrChildren.push({ tag: 'rFont', attrs: { val: run.font }, children: [] });
  if (run.color) rPrChildren.push({ tag: 'color', attrs: { rgb: run.color }, children: [] });

  const children: ParsedNode[] = [];
  if (rPrChildren.length > 0) {
    children.push({ tag: 'rPr', attrs: {}, children: rPrChildren });
  }
  children.push({ tag: 't', attrs: {}, children: [run.text] });

  return { tag: 'r', attrs: {}, children };
}
