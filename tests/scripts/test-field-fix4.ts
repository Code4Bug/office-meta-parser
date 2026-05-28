import { parseDocx } from '../../src/docx/index.js';
import * as fs from 'fs';

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { raw, semantic } = await parseDocx(buffer);

  // Print the full footer XML structure
  const footer = raw.parts.get('word/footer1.xml');
  if (footer) {
    const json = JSON.stringify(footer, null, 2);
    // Find PAGE context
    const idx = json.indexOf('PAGE');
    if (idx >= 0) {
      console.log('=== Footer XML around PAGE ===');
      console.log(json.substring(Math.max(0, idx - 500), Math.min(json.length, idx + 500)));
    }
  }

  // Check what the semantic parser produces for footer
  if (semantic.footers) {
    for (const f of semantic.footers) {
      console.log(`\n=== Footer ${f.id} semantic ===`);
      console.log(JSON.stringify(f, null, 2));
    }
  }
}

main();
