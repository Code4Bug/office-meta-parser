import { parseDocx } from '../../src/docx/index.js';
import { findChild } from '../../src/docx/parsers/utils.js';
import { parseParagraph } from '../../src/docx/parsers/paragraph.js';
import type { ParsedNode } from '../../src/core/types.js';
import type { Paragraph } from '../../src/docx/types.js';
import * as fs from 'fs';

function extractContentDebug(node: ParsedNode): Paragraph[] {
  const body = findChild(node, 'w:hdr') || findChild(node, 'w:ftr') || node;
  const paragraphs: Paragraph[] = [];

  let callCount = 0;
  function collectParagraphs(n: ParsedNode, depth: number) {
    callCount++;
    for (const child of n.children) {
      if (typeof child === 'string') continue;
      if (child.tag === 'w:p') {
        console.log(`Found w:p at depth ${depth} (call #${callCount})`);
        paragraphs.push(parseParagraph(child));
      }
      collectParagraphs(child, depth + 1);
    }
  }

  collectParagraphs(body, 0);
  console.log(`Total calls: ${callCount}`);
  return paragraphs;
}

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { raw } = await parseDocx(buffer);

  const footer = raw.parts.get('word/footer1.xml');
  if (footer) {
    const paras = extractContentDebug(footer);
    console.log(`\nTotal paragraphs: ${paras.length}`);
    for (let i = 0; i < paras.length; i++) {
      const para = paras[i];
      console.log(`\nParagraph ${i}: ${para.runs.length} runs`);
      for (let j = 0; j < para.runs.length; j++) {
        const run = para.runs[j];
        if (run.field) {
          console.log(`  run[${j}]: FIELD instr="${run.field.instruction}" result="${run.field.result}"`);
        } else {
          console.log(`  run[${j}]: text="${run.text}"`);
        }
      }
    }
  }
}

main();
