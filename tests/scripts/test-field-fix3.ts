import { parseDocx } from '../../src/docx/index.js';
import * as fs from 'fs';

async function main() {
  const buffer = fs.readFileSync('./tests/files/demo-00.docx').buffer;
  const { raw, semantic } = await parseDocx(buffer);

  // Check all XML parts for PAGE field
  for (const [path, node] of raw.parts) {
    if (path.includes('footer') || path.includes('header')) {
      const xml = JSON.stringify(node);
      if (xml.includes('PAGE') || xml.includes('NUMPAGES')) {
        console.log(`Found PAGE in: ${path}`);
        console.log(JSON.stringify(node, null, 2).substring(0, 2000));
      }
    }
  }

  // Also check document.xml for textboxes
  const docXml = raw.parts.get('word/document.xml');
  if (docXml) {
    const xml = JSON.stringify(docXml);
    if (xml.includes('PAGE') || xml.includes('NUMPAGES')) {
      console.log('\nFound PAGE in document.xml');
      // Find the path to PAGE
      const idx = xml.indexOf('PAGE');
      console.log('Context:', xml.substring(Math.max(0, idx - 200), idx + 200));
    }
  }
}

main();
