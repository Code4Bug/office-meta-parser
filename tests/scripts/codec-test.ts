#!/usr/bin/env tsx
/**
 * 编解码测试脚本
 * 加载测试文件，进行解析和序列化，将过程数据保存到 tests/logs
 */

import { writeFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { loadDocx, parseDocx, serializeDocx } from '../../src/docx/index.js';
import { loadXlsx, parseXlsx, serializeXlsx } from '../../src/xlsx/index.js';
import { loadPptx, parsePptx, serializePptx } from '../../src/pptx/index.js';
import { saveToFile } from '../../src/core/io.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const INPUT_DIR = join(__dirname, '..', 'input');
const LOGS_DIR = join(__dirname, '..', 'logs');

mkdirSync(LOGS_DIR, { recursive: true });

interface TestResult {
  format: string;
  fileName: string;
  success: boolean;
  parseTime: number;
  serializeTime: number;
  parseResult?: any;
  error?: string;
}

function saveJson(fileName: string, data: any): void {
  const filePath = join(LOGS_DIR, fileName);
  writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf-8');
  console.log(`  Saved: ${fileName}`);
}

async function testDocx(): Promise<TestResult> {
  console.log('\n=== Testing DOCX ===');
  const fileName = 'test.docx';
  const result: TestResult = {
    format: 'DOCX',
    fileName,
    success: false,
    parseTime: 0,
    serializeTime: 0,
  };

  try {
    // 1. 加载并解析
    console.log('1. Loading and parsing...');
    const startTime = performance.now();
    const { raw, semantic } = await loadDocx(join(INPUT_DIR, fileName));
    result.parseTime = performance.now() - startTime;

    console.log(`   Parse time: ${result.parseTime.toFixed(2)}ms`);
    console.log(`   Parts: ${raw.parts.size}`);
    console.log(`   Blocks: ${semantic.body.blocks.length}`);

    // 保存解析结果
    result.parseResult = {
      meta: semantic.meta,
      bodyBlocksCount: semantic.body.blocks.length,
      blocks: semantic.body.blocks.map((block: any, i: number) => ({
        index: i,
        type: block.type,
        content: block.type === 'paragraph'
          ? block.runs?.map((r: any) => r.text).join('').substring(0, 50)
          : block.type === 'table'
            ? `[${block.rows?.length} rows]`
            : block.type === 'hyperlink'
              ? block.runs?.[0]?.text
              : undefined,
      })),
    };
    saveJson('docx-parse-result.json', result.parseResult);

    // 保存原始语义数据（完整）
    saveJson('docx-semantic-full.json', semantic);

    // 2. 序列化
    console.log('2. Serializing...');
    const serializeStart = performance.now();
    const output = await serializeDocx(semantic);
    result.serializeTime = performance.now() - serializeStart;

    console.log(`   Serialize time: ${result.serializeTime.toFixed(2)}ms`);
    console.log(`   Output size: ${output.byteLength} bytes`);

    // 保存序列化结果
    await saveToFile(output, join(LOGS_DIR, 'docx-serialized.docx'));
    console.log('  Saved: docx-serialized.docx');

    // 3. 重新解析验证
    console.log('3. Re-parsing for verification...');
    const { semantic: restored } = await parseDocx(output);
    console.log(`   Restored blocks: ${restored.body.blocks.length}`);
    console.log(`   Match: ${restored.body.blocks.length === semantic.body.blocks.length}`);

    saveJson('docx-restored-semantic.json', restored);

    result.success = true;
    console.log('✓ DOCX test passed');
  } catch (error: any) {
    result.error = error.message;
    console.error('✗ DOCX test failed:', error.message);
  }

  return result;
}

async function testXlsx(): Promise<TestResult> {
  console.log('\n=== Testing XLSX ===');
  const fileName = 'test.xlsx';
  const result: TestResult = {
    format: 'XLSX',
    fileName,
    success: false,
    parseTime: 0,
    serializeTime: 0,
  };

  try {
    // 1. 加载并解析
    console.log('1. Loading and parsing...');
    const startTime = performance.now();
    const { raw, semantic } = await loadXlsx(join(INPUT_DIR, fileName));
    result.parseTime = performance.now() - startTime;

    console.log(`   Parse time: ${result.parseTime.toFixed(2)}ms`);
    console.log(`   Parts: ${raw.parts.size}`);
    console.log(`   Sheets: ${semantic.sheets.length}`);
    console.log(`   Shared strings: ${semantic.sharedStrings.length}`);

    // 保存解析结果
    result.parseResult = {
      meta: semantic.meta,
      sheetsCount: semantic.sheets.length,
      sharedStringsCount: semantic.sharedStrings.length,
      sheets: semantic.sheets.map((sheet: any, i: number) => ({
        index: i,
        name: sheet.name,
        cellsCount: sheet.cells.length,
        hyperlinksCount: sheet.hyperlinks?.length || 0,
        hasAutoFilter: !!sheet.autoFilter,
        hasDataValidations: !!sheet.dataValidations?.length,
        hasConditionalFormats: !!sheet.conditionalFormats?.length,
        hasPrintArea: !!sheet.printArea,
      })),
    };
    saveJson('xlsx-parse-result.json', result.parseResult);

    // 保存完整语义数据
    saveJson('xlsx-semantic-full.json', semantic);

    // 2. 序列化
    console.log('2. Serializing...');
    const serializeStart = performance.now();
    const output = await serializeXlsx(semantic);
    result.serializeTime = performance.now() - serializeStart;

    console.log(`   Serialize time: ${result.serializeTime.toFixed(2)}ms`);
    console.log(`   Output size: ${output.byteLength} bytes`);

    // 保存序列化结果
    await saveToFile(output, join(LOGS_DIR, 'xlsx-serialized.xlsx'));
    console.log('  Saved: xlsx-serialized.xlsx');

    // 3. 重新解析验证
    console.log('3. Re-parsing for verification...');
    const { semantic: restored } = await parseXlsx(output);
    console.log(`   Restored sheets: ${restored.sheets.length}`);
    console.log(`   Match: ${restored.sheets.length === semantic.sheets.length}`);

    saveJson('xlsx-restored-semantic.json', restored);

    result.success = true;
    console.log('✓ XLSX test passed');
  } catch (error: any) {
    result.error = error.message;
    console.error('✗ XLSX test failed:', error.message);
  }

  return result;
}

async function testPptx(): Promise<TestResult> {
  console.log('\n=== Testing PPTX ===');
  const fileName = 'test.pptx';
  const result: TestResult = {
    format: 'PPTX',
    fileName,
    success: false,
    parseTime: 0,
    serializeTime: 0,
  };

  try {
    // 1. 加载并解析
    console.log('1. Loading and parsing...');
    const startTime = performance.now();
    const { raw, semantic } = await loadPptx(join(INPUT_DIR, fileName));
    result.parseTime = performance.now() - startTime;

    console.log(`   Parse time: ${result.parseTime.toFixed(2)}ms`);
    console.log(`   Parts: ${raw.parts.size}`);
    console.log(`   Slides: ${semantic.slides.length}`);

    // 保存解析结果
    result.parseResult = {
      meta: semantic.meta,
      slidesCount: semantic.slides.length,
      slides: semantic.slides.map((slide: any, i: number) => ({
        index: i,
        elementsCount: slide.elements.length,
        elementTypes: slide.elements.map((e: any) => e.type),
        hasTransition: !!slide.transition,
        hasAnimations: !!slide.animations?.length,
        hasNotes: !!slide.notes,
      })),
    };
    saveJson('pptx-parse-result.json', result.parseResult);

    // 保存完整语义数据
    saveJson('pptx-semantic-full.json', semantic);

    // 2. 序列化
    console.log('2. Serializing...');
    const serializeStart = performance.now();
    const output = await serializePptx(semantic);
    result.serializeTime = performance.now() - serializeStart;

    console.log(`   Serialize time: ${result.serializeTime.toFixed(2)}ms`);
    console.log(`   Output size: ${output.byteLength} bytes`);

    // 保存序列化结果
    await saveToFile(output, join(LOGS_DIR, 'pptx-serialized.pptx'));
    console.log('  Saved: pptx-serialized.pptx');

    // 3. 重新解析验证
    console.log('3. Re-parsing for verification...');
    const { semantic: restored } = await parsePptx(output);
    console.log(`   Restored slides: ${restored.slides.length}`);
    console.log(`   Match: ${restored.slides.length === semantic.slides.length}`);

    saveJson('pptx-restored-semantic.json', restored);

    result.success = true;
    console.log('✓ PPTX test passed');
  } catch (error: any) {
    result.error = error.message;
    console.error('✗ PPTX test failed:', error.message);
  }

  return result;
}

async function main() {
  console.log('='.repeat(60));
  console.log('Office Meta Parser - Codec Test');
  console.log('='.repeat(60));
  console.log(`Logs directory: ${LOGS_DIR}`);
  console.log(`Timestamp: ${new Date().toISOString()}`);

  const results: TestResult[] = [];

  // 运行所有测试
  results.push(await testDocx());
  results.push(await testXlsx());
  results.push(await testPptx());

  // 生成总结报告
  console.log('\n' + '='.repeat(60));
  console.log('Summary');
  console.log('='.repeat(60));

  const summary = {
    timestamp: new Date().toISOString(),
    totalTests: results.length,
    passed: results.filter(r => r.success).length,
    failed: results.filter(r => !r.success).length,
    results: results.map(r => ({
      format: r.format,
      fileName: r.fileName,
      success: r.success,
      parseTime: `${r.parseTime.toFixed(2)}ms`,
      serializeTime: `${r.serializeTime.toFixed(2)}ms`,
      error: r.error,
    })),
  };

  for (const r of results) {
    const status = r.success ? '✓' : '✗';
    console.log(`${status} ${r.format}: ${r.fileName}`);
    if (r.success) {
      console.log(`  Parse: ${r.parseTime.toFixed(2)}ms, Serialize: ${r.serializeTime.toFixed(2)}ms`);
    } else {
      console.log(`  Error: ${r.error}`);
    }
  }

  // 保存总结报告
  saveJson('test-summary.json', summary);

  console.log('\n' + '='.repeat(60));
  console.log(`All logs saved to: ${LOGS_DIR}`);
  console.log('='.repeat(60));

  // 如果有失败的测试，退出码为 1
  if (summary.failed > 0) {
    process.exit(1);
  }
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
