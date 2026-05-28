import { parseDocx } from '../../src/docx/index.js';
import * as fs from 'fs';

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { raw } = await parseDocx(buffer);

  const footer = raw.parts.get('word/footer1.xml');
  if (footer) {
    console.log(JSON.stringify(footer, null, 2));
  }
}

main();
