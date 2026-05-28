import { readdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { parseXml } from '../../src/core/xml.js';
import { unzip } from '../../src/core/zip.js';
import { loadFromFile } from '../../src/core/io.js';
import type { ParsedNode } from '../../src/core/types.js';

const __dirname = dirname(fileURLToPath(import.meta.url));

// 递归检查 XML 结构问题
function findIssues(node: ParsedNode, path: string, issues: string[]) {
  const currentPath = `${path}/${node.tag}`;

  // 1. spPr 下 a:solidFill 必须在 a:xfrm 之后、a:prstGeom 之后
  // 2. p:sp 必须有 p:nvSpPr
  // 3. p:txBody 必须有至少一个 a:p
  // 4. a:ln 必须有子元素或 w 属性

  if (node.tag === 'p:txBody') {
    const hasP = node.children.some(c => typeof c !== 'string' && c.tag === 'a:p');
    if (!hasP) {
      issues.push(`${currentPath}: txBody missing <a:p>`);
    }
  }

  if (node.tag === 'p:sp' || node.tag === 'p:grpSp' || node.tag === 'p:pic' || node.tag === 'p:graphicFrame') {
    const hasNvPr = node.children.some(c =>
      typeof c !== 'string' && (c.tag === 'p:nvSpPr' || c.tag === 'p:nvGrpSpPr' || c.tag === 'p:nvPicPr' || c.tag === 'p:nvGraphicFramePr')
    );
    if (!hasNvPr) {
      issues.push(`${currentPath}: missing nvXxxPr`);
    }
  }

  if (node.tag === 'a:ln') {
    if (node.children.length === 0 && !node.attrs.w) {
      issues.push(`${currentPath}: empty <a:ln> with no w attr and no children`);
    }
  }

  // 检查 spPr 中 solidFill 的位置
  if (node.tag === 'p:spPr' || node.tag === 'p:grpSpPr') {
    const childTags = node.children.filter(c => typeof c !== 'string').map(c => (c as ParsedNode).tag);
    const xfrmIdx = childTags.indexOf('a:xfrm');
    const fillIdx = childTags.indexOf('a:solidFill');
    if (fillIdx !== -1 && xfrmIdx !== -1 && fillIdx < xfrmIdx) {
      issues.push(`${currentPath}: a:solidFill before a:xfrm (wrong order)`);
    }
  }

  for (const child of node.children) {
    if (typeof child !== 'string') {
      findIssues(child, currentPath, issues);
    }
  }
}

async function main() {
  // 先检查真实文件作为基线
  const realAb = await loadFromFile(join(__dirname, '..', 'input', 'test.pptx'));
  const realEntries = await unzip(realAb);

  console.log('=== Real file validation ===');
  for (const entry of realEntries) {
    if (entry.path.endsWith('.xml') && !entry.path.endsWith('.rels')) {
      const xml = new TextDecoder().decode(entry.data);
      const root = parseXml(xml);
      const issues: string[] = [];
      findIssues(root, '', issues);
      if (issues.length > 0) {
        console.log(`  ${entry.path}:`);
        for (const issue of issues) console.log(`    ${issue}`);
      }
    }
  }
  console.log('  (done)\n');

  // 检查生成的文件
  const dir = join(__dirname, '..', 'fixtures', 'generated');
  const files = readdirSync(dir).filter(f => f.endsWith('.pptx'));

  for (const file of files) {
    const ab = await loadFromFile(join(dir, file));
    const entries = await unzip(ab);
    const issues: string[] = [];

    for (const entry of entries) {
      if (entry.path.endsWith('.xml') && !entry.path.endsWith('.rels')) {
        const xml = new TextDecoder().decode(entry.data);
        const root = parseXml(xml);
        findIssues(root, entry.path, issues);
      }
    }

    if (issues.length > 0) {
      console.log(`${file}:`);
      for (const issue of issues) console.log(`  ${issue}`);
    }
  }
}
main();
