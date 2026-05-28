import type { ParsedNode } from '../../core/types.js';
import type { Paragraph, ParagraphProperties, TextRun, NumberingProperties } from '../types.js';

export function serializeParagraph(para: Paragraph): ParsedNode {
  const children: ParsedNode[] = [];

  if (para.style || para.properties || para.numbering) {
    children.push(serializeParagraphProperties(para.properties, para.numbering, para.style));
  }

  // 按修订标记分组: 连续的相同修订 run 合并到一个 w:ins/w:del 中
  let i = 0;
  while (i < para.runs.length) {
    const run = para.runs[i];

    if (run.field) {
      children.push(...serializeField(run));
      i++;
      continue;
    }

    if (run.revisionType) {
      // 收集连续的相同修订 run
      const group: TextRun[] = [run];
      let j = i + 1;
      while (j < para.runs.length) {
        const next = para.runs[j];
        if (next.revisionType === run.revisionType
          && next.revisionAuthor === run.revisionAuthor
          && next.revisionDate === run.revisionDate
          && !next.field) {
          group.push(next);
          j++;
        } else {
          break;
        }
      }

      const tag = run.revisionType === 'insert' ? 'w:ins' : 'w:del';
      const revChildren = group.map(r => serializeRun(r, run.revisionType === 'delete'));
      children.push({
        tag,
        attrs: {
          'w:author': run.revisionAuthor || '',
          'w:date': run.revisionDate || '',
        },
        children: revChildren,
      });
      i = j;
    } else {
      if (run.commentId) {
        children.push({ tag: 'w:commentRangeStart', attrs: { 'w:id': run.commentId }, children: [] });
        children.push(serializeRun(run));
        children.push({ tag: 'w:commentRangeEnd', attrs: { 'w:id': run.commentId }, children: [] });
        children.push({
          tag: 'w:r',
          attrs: {},
          children: [{ tag: 'w:commentReference', attrs: { 'w:id': run.commentId }, children: [] }],
        });
      } else {
        children.push(serializeRun(run));
      }
      i++;
    }
  }

  return {
    tag: 'w:p',
    attrs: {},
    children,
  };
}

function serializeField(run: TextRun): ParsedNode[] {
  const result: ParsedNode[] = [];
  const rPr = serializeRunProperties(run);

  // begin
  const beginChildren: ParsedNode[] = [];
  if (rPr) beginChildren.push(rPr);
  beginChildren.push({ tag: 'w:fldChar', attrs: { 'w:fldCharType': 'begin' }, children: [] });
  result.push({ tag: 'w:r', attrs: {}, children: beginChildren });

  // instruction
  const instrChildren: ParsedNode[] = [];
  if (rPr) instrChildren.push(rPr);
  instrChildren.push({
    tag: 'w:instrText',
    attrs: { 'xml:space': 'preserve' },
    children: [' ' + run.field!.instruction + ' '],
  });
  result.push({ tag: 'w:r', attrs: {}, children: instrChildren });

  // separate
  const sepChildren: ParsedNode[] = [];
  if (rPr) sepChildren.push(rPr);
  sepChildren.push({ tag: 'w:fldChar', attrs: { 'w:fldCharType': 'separate' }, children: [] });
  result.push({ tag: 'w:r', attrs: {}, children: sepChildren });

  // result
  if (run.field!.result) {
    result.push(serializeRun({ ...run, text: run.field!.result, field: undefined }));
  }

  // end
  const endChildren: ParsedNode[] = [];
  if (rPr) endChildren.push(rPr);
  endChildren.push({ tag: 'w:fldChar', attrs: { 'w:fldCharType': 'end' }, children: [] });
  result.push({ tag: 'w:r', attrs: {}, children: endChildren });

  return result;
}

export function serializeRun(run: TextRun, useDelText = false): ParsedNode {
  const children: (ParsedNode | string)[] = [];

  const rPr = serializeRunProperties(run);
  if (rPr) {
    children.push(rPr);
  }

  children.push({
    tag: useDelText ? 'w:delText' : 'w:t',
    attrs: { 'xml:space': 'preserve' },
    children: [run.text],
  });

  return {
    tag: 'w:r',
    attrs: {},
    children,
  };
}

export function serializeParagraphProperties(props?: ParagraphProperties, numbering?: NumberingProperties, style?: string): ParsedNode {
  const children: ParsedNode[] = [];

  if (style) {
    children.push({ tag: 'w:pStyle', attrs: { 'w:val': style }, children: [] });
  }

  if (props?.alignment) {
    const val = props.alignment === 'justify' ? 'both' : props.alignment;
    children.push({ tag: 'w:jc', attrs: { 'w:val': val }, children: [] });
  }

  if (numbering) {
    const numPrChildren: ParsedNode[] = [
      { tag: 'w:ilvl', attrs: { 'w:val': String(numbering.level) }, children: [] },
      { tag: 'w:numId', attrs: { 'w:val': numbering.numId }, children: [] },
    ];
    children.push({ tag: 'w:numPr', attrs: {}, children: numPrChildren });
  }

  if (props?.indent) {
    const attrs: Record<string, string> = {};
    if (props.indent.left !== undefined) attrs['w:left'] = String(props.indent.left);
    if (props.indent.right !== undefined) attrs['w:right'] = String(props.indent.right);
    if (props.indent.start !== undefined) attrs['w:start'] = String(props.indent.start);
    if (props.indent.end !== undefined) attrs['w:end'] = String(props.indent.end);
    if (props.indent.firstLine !== undefined) attrs['w:firstLine'] = String(props.indent.firstLine);
    if (props.indent.hanging !== undefined) attrs['w:hanging'] = String(props.indent.hanging);
    if (props.indent.leftChars !== undefined) attrs['w:leftChars'] = String(props.indent.leftChars);
    if (props.indent.rightChars !== undefined) attrs['w:rightChars'] = String(props.indent.rightChars);
    if (props.indent.startChars !== undefined) attrs['w:startChars'] = String(props.indent.startChars);
    if (props.indent.endChars !== undefined) attrs['w:endChars'] = String(props.indent.endChars);
    if (props.indent.firstLineChars !== undefined) attrs['w:firstLineChars'] = String(props.indent.firstLineChars);
    if (props.indent.hangingChars !== undefined) attrs['w:hangingChars'] = String(props.indent.hangingChars);
    if (Object.keys(attrs).length > 0) {
      children.push({ tag: 'w:ind', attrs, children: [] });
    }
  }

  if (props?.spacing) {
    const attrs: Record<string, string> = {};
    if (props.spacing.before !== undefined) attrs['w:before'] = String(props.spacing.before);
    if (props.spacing.after !== undefined) attrs['w:after'] = String(props.spacing.after);
    if (props.spacing.line !== undefined) attrs['w:line'] = String(props.spacing.line);
    if (props.spacing.lineRule) attrs['w:lineRule'] = props.spacing.lineRule;
    if (Object.keys(attrs).length > 0) {
      children.push({ tag: 'w:spacing', attrs, children: [] });
    }
  }

  if (props?.keepNext) {
    children.push({ tag: 'w:keepNext', attrs: {}, children: [] });
  }
  if (props?.keepLines) {
    children.push({ tag: 'w:keepLines', attrs: {}, children: [] });
  }
  if (props?.pageBreakBefore) {
    children.push({ tag: 'w:pageBreakBefore', attrs: {}, children: [] });
  }
  if (props?.outlineLevel !== undefined) {
    children.push({ tag: 'w:outlineLvl', attrs: { 'w:val': String(props.outlineLevel) }, children: [] });
  }
  if (props?.shading) {
    const attrs: Record<string, string> = {};
    if (props.shading.pattern) attrs['w:val'] = props.shading.pattern;
    if (props.shading.fill) attrs['w:fill'] = props.shading.fill;
    if (props.shading.color) attrs['w:color'] = props.shading.color;
    children.push({ tag: 'w:shd', attrs, children: [] });
  }
  if (props?.border) {
    const bdrChildren: ParsedNode[] = [];
    const borderTags = ['top', 'left', 'bottom', 'right', 'between'] as const;
    for (const side of borderTags) {
      const b = props.border[side];
      if (b) {
        const attrs: Record<string, string> = { 'w:val': b.style };
        if (b.size !== undefined) attrs['w:sz'] = String(b.size);
        if (b.color) attrs['w:color'] = b.color;
        bdrChildren.push({ tag: `w:${side}`, attrs, children: [] });
      }
    }
    if (bdrChildren.length > 0) {
      children.push({ tag: 'w:pBdr', attrs: {}, children: bdrChildren });
    }
  }
  if (props?.tabs && props.tabs.length > 0) {
    const tabChildren: ParsedNode[] = props.tabs.map(tab => ({
      tag: 'w:tab',
      attrs: {
        'w:val': tab.alignment,
        'w:pos': String(tab.position),
        ...(tab.leader ? { 'w:leader': tab.leader } : {}),
      },
      children: [],
    }));
    children.push({ tag: 'w:tabs', attrs: {}, children: tabChildren });
  }

  if (props?.formatChangeAuthor) {
    children.push({
      tag: 'w:pPrChange',
      attrs: {
        'w:author': props.formatChangeAuthor,
        'w:date': props.formatChangeDate || '',
      },
      children: [],
    });
  }

  return { tag: 'w:pPr', attrs: {}, children };
}

export function serializeRunProperties(run: TextRun): ParsedNode | null {
  const children: ParsedNode[] = [];

  if (run.bold) {
    children.push({ tag: 'w:b', attrs: {}, children: [] });
  }
  if (run.italic) {
    children.push({ tag: 'w:i', attrs: {}, children: [] });
  }
  if (run.underline) {
    children.push({ tag: 'w:u', attrs: { 'w:val': 'single' }, children: [] });
  }
  if (run.strike) {
    children.push({ tag: 'w:strike', attrs: {}, children: [] });
  }
  if (run.fontSize) {
    children.push({ tag: 'w:sz', attrs: { 'w:val': String(run.fontSize) }, children: [] });
  }
  if (run.color) {
    children.push({ tag: 'w:color', attrs: { 'w:val': run.color }, children: [] });
  }
  if (run.fontFamily) {
    children.push({
      tag: 'w:rFonts',
      attrs: { 'w:ascii': run.fontFamily, 'w:hAnsi': run.fontFamily },
      children: [],
    });
  }
  if (run.superscript) {
    children.push({ tag: 'w:vertAlign', attrs: { 'w:val': 'superscript' }, children: [] });
  }
  if (run.subscript) {
    children.push({ tag: 'w:vertAlign', attrs: { 'w:val': 'subscript' }, children: [] });
  }
  if (run.highlight) {
    children.push({ tag: 'w:highlight', attrs: { 'w:val': run.highlight }, children: [] });
  }
  if (run.shadingColor !== undefined || run.shadingPattern !== undefined) {
    const attrs: Record<string, string> = {};
    if (run.shadingPattern) attrs['w:val'] = run.shadingPattern;
    if (run.shadingColor) attrs['w:fill'] = run.shadingColor;
    children.push({ tag: 'w:shd', attrs, children: [] });
  }
  if (run.caps) {
    children.push({ tag: 'w:caps', attrs: {}, children: [] });
  }
  if (run.smallCaps) {
    children.push({ tag: 'w:smallCaps', attrs: {}, children: [] });
  }
  if (run.dstrike) {
    children.push({ tag: 'w:dstrike', attrs: {}, children: [] });
  }
  if (run.vanish) {
    children.push({ tag: 'w:vanish', attrs: {}, children: [] });
  }
  if (run.characterSpacing !== undefined) {
    children.push({ tag: 'w:spacing', attrs: { 'w:val': String(run.characterSpacing) }, children: [] });
  }
  if (run.kern !== undefined) {
    children.push({ tag: 'w:kern', attrs: { 'w:val': String(run.kern) }, children: [] });
  }
  if (run.emphasis) {
    children.push({ tag: 'w:em', attrs: { 'w:val': run.emphasis }, children: [] });
  }
  if (run.characterBorder) {
    const bdrAttrs: Record<string, string> = {
      'w:val': run.characterBorder.style,
      'w:space': '0',
      'w:color': run.characterBorder.color || 'auto',
    };
    if (run.characterBorder.size !== undefined) bdrAttrs['w:sz'] = String(run.characterBorder.size);
    children.push({ tag: 'w:bdr', attrs: bdrAttrs, children: [] });
  }
  if (run.verticalPosition !== undefined) {
    children.push({ tag: 'w:position', attrs: { 'w:val': String(run.verticalPosition) }, children: [] });
  }
  if (run.properties?.formatChangeAuthor) {
    children.push({
      tag: 'w:rPrChange',
      attrs: {
        'w:author': run.properties.formatChangeAuthor,
        'w:date': run.properties.formatChangeDate || '',
      },
      children: [],
    });
  }

  if (children.length === 0) return null;

  return {
    tag: 'w:rPr',
    attrs: {},
    children,
  };
}
