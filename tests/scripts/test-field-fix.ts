import { parseDocx } from '../../src/docx/index.js';
import * as fs from 'fs';

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { semantic } = await parseDocx(buffer);

  console.log('=== demo-00.docx 正文段落 ===');
  for (const block of semantic.body.blocks) {
    if (block.type === 'paragraph') {
      const parts: string[] = [];
      for (const run of block.runs) {
        if (run.field) {
          parts.push(`[FIELD: instr="${run.field.instruction}", result="${run.field.result}"]`);
        } else {
          parts.push(run.text);
        }
      }
      const line = parts.join('');
      if (line.includes('PAGE') || line.includes('NUMPAGES') || line.includes('第') || line.includes('页')) {
        console.log(`  "${line}"`);
        console.log('  runs:', JSON.stringify(block.runs.map(r => r.field ? { field: r.field } : { text: r.text }), null, 2));
      }
    }
  }
}

main();
