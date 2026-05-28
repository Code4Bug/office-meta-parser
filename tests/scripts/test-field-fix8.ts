import { parseDocx } from '../../src/docx/index.js';
import { findChild } from '../../src/docx/parsers/utils.js';
import type { ParsedNode } from '../../src/core/types.js';
import * as fs from 'fs';

function dumpTree(n: ParsedNode, depth: number, maxDepth: number) {
  if (depth > maxDepth) return;
  for (const child of n.children) {
    if (typeof child === 'string') {
      if (child.trim()) {
        console.log(`${'  '.repeat(depth)}[text: "${child.trim().substring(0, 50)}"]`);
      }
    } else {
      console.log(`${'  '.repeat(depth)}<${child.tag}> (${child.children.length} children)`);
      dumpTree(child, depth + 1, maxDepth);
    }
  }
}

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { raw } = await parseDocx(buffer);

  const footer = raw.parts.get('word/footer1.xml');
  if (footer) {
    const body = findChild(footer, 'w:ftr') || footer;
    console.log('Body tag:', (body as any).tag);
    console.log('Body children count:', body.children.length);
    dumpTree(body, 0, 5);
  }
}

main();
