import { parseDocx } from '../../src/docx/index.js';
import * as fs from 'fs';

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { semantic } = await parseDocx(buffer);

  console.log('=== Headers ===');
  if (semantic.headers) {
    for (const header of semantic.headers) {
      console.log(`Header ${header.id}:`);
      for (const para of header.content) {
        const parts: string[] = [];
        for (const run of para.runs) {
          if (run.field) {
            parts.push(`[FIELD: instr="${run.field.instruction}", result="${run.field.result}"]`);
          } else {
            parts.push(run.text);
          }
        }
        console.log(`  "${parts.join('')}"`);
      }
    }
  }

  console.log('\n=== Footers ===');
  if (semantic.footers) {
    for (const footer of semantic.footers) {
      console.log(`Footer ${footer.id}:`);
      for (const para of footer.content) {
        const parts: string[] = [];
        for (const run of para.runs) {
          if (run.field) {
            parts.push(`[FIELD: instr="${run.field.instruction}", result="${run.field.result}"]`);
          } else {
            parts.push(run.text);
          }
        }
        console.log(`  "${parts.join('')}"`);
      }
    }
  }
}

main();
