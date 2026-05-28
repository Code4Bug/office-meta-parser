import { parseDocx } from '../../src/docx/index.js';
import * as fs from 'fs';

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { raw, semantic } = await parseDocx(buffer);

  if (semantic.footers) {
    for (const f of semantic.footers) {
      console.log(`Footer ${f.id} (${f.type}):`);
      console.log(`  paragraphs: ${f.content.length}`);
      for (let i = 0; i < f.content.length; i++) {
        const para = f.content[i];
        console.log(`  [${i}] runs: ${para.runs.length}`);
        for (let j = 0; j < para.runs.length; j++) {
          const run = para.runs[j];
          if (run.field) {
            console.log(`    run[${j}]: FIELD instr="${run.field.instruction}" result="${run.field.result}"`);
          } else {
            console.log(`    run[${j}]: text="${run.text}" bold=${run.bold} fontSize=${run.fontSize}`);
          }
        }
      }
    }
  }
}

main();
