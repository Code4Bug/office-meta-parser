import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { RichTextRun, SharedStringEntry } from '../types.js';
import { findChild, getTextContent } from './utils.js';

export function extractSharedStrings(raw: RawDocument): SharedStringEntry[] {
  const sstXml = raw.parts.get('xl/sharedStrings.xml');
  if (!sstXml) return [];

  const strings: SharedStringEntry[] = [];
  for (const child of sstXml.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'si') continue;
    const text = extractSiText(child);
    const richText = extractSiRichText(child);
    strings.push(richText ? { text, richText } : { text });
  }
  return strings;
}

export function extractSharedStringRichText(raw: RawDocument, index: number): RichTextRun[] | undefined {
  const sstXml = raw.parts.get('xl/sharedStrings.xml');
  if (!sstXml) return undefined;

  let i = 0;
  for (const child of sstXml.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'si') continue;
    if (i === index) return extractSiRichText(child);
    i++;
  }
  return undefined;
}

function extractSiText(node: ParsedNode): string {
  const tNode = findChild(node, 't');
  if (tNode) return getTextContent(tNode);

  // Rich text: concatenate all <t> inside <r> elements
  let text = '';
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'r') {
      const t = findChild(child, 't');
      if (t) text += getTextContent(t);
    }
  }
  return text;
}

function extractSiRichText(node: ParsedNode): RichTextRun[] | undefined {
  const hasRichText = node.children.some(
    c => typeof c !== 'string' && c.tag === 'r' && findChild(c, 'rPr'),
  );
  if (!hasRichText) return undefined;

  const runs: RichTextRun[] = [];
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'r') {
      const rPr = findChild(child, 'rPr');
      const t = findChild(child, 't');
      const text = t ? getTextContent(t) : '';
      if (!rPr) {
        runs.push({ text });
      } else {
        const run: RichTextRun = { text };
        for (const prop of rPr.children) {
          if (typeof prop === 'string') continue;
          switch (prop.tag) {
            case 'b': run.bold = true; break;
            case 'i': run.italic = true; break;
            case 'u': run.underline = true; break;
            case 'strike': run.strike = true; break;
            case 'sz': run.size = parseFloat(prop.attrs['val'] || '0'); break;
            case 'rFont': run.font = prop.attrs['val']; break;
            case 'color': {
              const rgb = prop.attrs['rgb'];
              if (rgb) run.color = rgb;
              break;
            }
          }
        }
        runs.push(run);
      }
    }
  }
  return runs;
}
