import { describe, it, expect } from 'vitest';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { loadDocx, parseDocx, serializeDocx } from '../../src/docx/index.js';
import { loadXlsx, parseXlsx, serializeXlsx } from '../../src/xlsx/index.js';
import { loadPptx, parsePptx, serializePptx } from '../../src/pptx/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const INPUT_DIR = join(__dirname, '..', 'input');

// ============================================================
// DOCX 测试
// ============================================================
describe('Real DOCX file', () => {
  it('parses docx meta correctly', async () => {
    const { raw, semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    expect(raw.parts.size).toBeGreaterThan(0);
    expect(semantic.meta.title).toBe('Office Meta Parser 测试文档');
    expect(semantic.meta.creator).toBe('张三');
    expect(semantic.meta.lastModifiedBy).toBe('李四');
    expect(semantic.meta.category).toBe('测试文档');
  });

  it('parses title paragraph with formatting', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    expect(semantic.body.blocks.length).toBeGreaterThanOrEqual(4);

    // 第一段是标题，包含粗体、大字号、颜色
    const heading = semantic.body.blocks[0] as any;
    expect(heading.type).toBe('paragraph');
    expect(heading.runs[0].text).toBe('Office Meta Parser 测试报告');
    expect(heading.runs[0].bold).toBe(true);
    expect(heading.runs[0].fontSize).toBe(56); // w:sz val="56" (half-points)
    expect(heading.runs[0].color).toBe('1F4E79');
  });

  it('parses mixed formatting (bold/italic/underline/strike/superscript/subscript)', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    // 第5段有混合格式
    const mixed = semantic.body.blocks[4] as any;
    expect(mixed.type).toBe('paragraph');
    // 应该有多个 runs: 文本格式测试：、粗体、顿号、斜体、...
    expect(mixed.runs.length).toBeGreaterThanOrEqual(9);

    // 找到各格式的 run
    const boldRun = mixed.runs.find((r: any) => r.text === '粗体');
    expect(boldRun?.bold).toBe(true);

    const italicRun = mixed.runs.find((r: any) => r.text === '斜体');
    expect(italicRun?.italic).toBe(true);

    const underlineRun = mixed.runs.find((r: any) => r.text === '下划线');
    expect(underlineRun?.underline).toBe(true);

    const strikeRun = mixed.runs.find((r: any) => r.text === '删除线');
    expect(strikeRun?.strike).toBe(true);

    const superscriptRun = mixed.runs.find((r: any) => r.text === '上标');
    expect(superscriptRun?.superscript).toBe(true);

    const subscriptRun = mixed.runs.find((r: any) => r.text === '下标');
    expect(subscriptRun?.subscript).toBe(true);
  });

  it('parses font size and color variations', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    // 第6段有不同字号和颜色
    const colorParagraph = semantic.body.blocks[5] as any;
    expect(colorParagraph.type).toBe('paragraph');

    const redSmall = colorParagraph.runs.find((r: any) => r.text === '红色小字');
    expect(redSmall?.fontSize).toBe(20); // w:sz val="20" (half-points)
    expect(redSmall?.color).toBe('FF0000');

    const blueMedium = colorParagraph.runs.find((r: any) => r.text === '蓝色中字');
    expect(blueMedium?.fontSize).toBe(36); // w:sz val="36" (half-points)
    expect(blueMedium?.color).toBe('0000FF');

    const greenLarge = colorParagraph.runs.find((r: any) => r.text === '绿色大字');
    expect(greenLarge?.fontSize).toBe(48); // w:sz val="48" (half-points)
    expect(greenLarge?.color).toBe('00FF00');
  });

  it('parses paragraph alignment', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    // 第7-10段是不同对齐方式
    const leftPara = semantic.body.blocks[6] as any;
    expect(leftPara.runs[0].text).toBe('左对齐文本');

    const centerPara = semantic.body.blocks[7] as any;
    expect(centerPara.runs[0].text).toBe('居中对齐文本');

    const rightPara = semantic.body.blocks[8] as any;
    expect(rightPara.runs[0].text).toBe('右对齐文本');

    const bothPara = semantic.body.blocks[9] as any;
    expect(bothPara.runs[0].text).toContain('两端对齐文本');
  });

  it('parses hyperlinks', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    // 查找超链接块（w:hyperlink 被解析为独立的 hyperlink 类型）
    const hyperlinkBlock = semantic.body.blocks.find((b: any) => b.type === 'hyperlink');
    expect(hyperlinkBlock).toBeDefined();
    expect((hyperlinkBlock as any).runs[0].text).toBe('GitHub');
  });

  it('parses table with merge (gridSpan/vMerge)', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    // 查找表格
    const tableBlock = semantic.body.blocks.find((b: any) => b.type === 'table');
    expect(tableBlock).toBeDefined();
    const table = tableBlock as any;
    expect(table.rows.length).toBeGreaterThanOrEqual(3);
    // 第一行应该有合并单元格
    expect(table.rows[0].cells.length).toBeGreaterThanOrEqual(3);
  });

  it('parses table with vertical alignment', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    // 查找第二个表格（垂直对齐测试）
    const tables = semantic.body.blocks.filter((b: any) => b.type === 'table');
    expect(tables.length).toBeGreaterThanOrEqual(2);
  });

  it('parses numbered list', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    // 查找编号列表项（第一项、第二项、第三项）
    const listItems = semantic.body.blocks.filter((b: any) =>
      b.type === 'paragraph' && b.runs?.some((r: any) =>
        r.text === '第一项' || r.text === '第二项' || r.text === '第三项'
      )
    );
    expect(listItems.length).toBe(3);
  });

  it('round-trips DOCX body', async () => {
    const { semantic } = await loadDocx(join(INPUT_DIR, 'test.docx'));

    const output = await serializeDocx(semantic);
    expect(output).toBeInstanceOf(ArrayBuffer);

    // 重新解析，body 应该被保留
    const { semantic: restored } = await parseDocx(output);
    expect(restored.body.blocks.length).toBe(semantic.body.blocks.length);
  });
});

// ============================================================
// XLSX 测试
// ============================================================
describe('Real XLSX file', () => {
  it('parses xlsx meta correctly', async () => {
    const { raw, semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    expect(raw.parts.size).toBeGreaterThan(0);
    expect(semantic.meta.title).toBe('销售数据分析表');
    expect(semantic.meta.creator).toBe('王五');
  });

  it('parses multiple sheets', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    expect(semantic.sheets).toHaveLength(3);
    expect(semantic.sheets[0].name).toBe('销售数据');
    expect(semantic.sheets[1].name).toBe('汇总统计');
    expect(semantic.sheets[2].name).toBe('格式测试');
  });

  it('parses shared strings', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const texts = semantic.sharedStrings.map(s => s.text);
    expect(texts).toContain('产品名称');
    expect(texts).toContain('笔记本电脑 Pro');
    expect(texts).toContain('合计');
    expect(texts).toContain('平均值');
  });

  it('parses cells with shared strings and numbers', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet1 = semantic.sheets[0];
    expect(sheet1.cells).toHaveLength(10); // rows 1-6 + 8-11 (row 7 is empty, not stored)

    // 表头行
    expect(sheet1.cells[0]).toHaveLength(7); // 7 columns
    expect(sheet1.cells[0][0].type).toBe('sharedString');
    expect(sheet1.cells[0][0].value).toBe('产品编号');

    // 数据行 - 数字
    expect(sheet1.cells[1][3].type).toBe('number');
    expect(sheet1.cells[1][3].value).toBe(8999);

    expect(sheet1.cells[1][4].type).toBe('number');
    expect(sheet1.cells[1][4].value).toBe(5);
  });

  it('parses inline strings', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet2 = semantic.sheets[1];
    expect(sheet2.cells[0][0].type).toBe('string');
    expect(sheet2.cells[0][0].value).toBe('类别统计');

    expect(sheet2.cells[2][0].type).toBe('string');
    expect(sheet2.cells[2][0].value).toBe('电脑');
  });

  it('parses formulas', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet1 = semantic.sheets[0];

    // 汇总行 - row 7 is empty, so row 8 is at index 6
    const sumCell = sheet1.cells[6]?.find((c: any) => c.formula);
    expect(sumCell?.formula).toBe('SUM(F2:F6)');

    const avgCell = sheet1.cells[7]?.find((c: any) => c.formula);
    expect(avgCell?.formula).toBe('AVERAGE(F2:F6)');

    const maxCell = sheet1.cells[8]?.find((c: any) => c.formula);
    expect(maxCell?.formula).toBe('MAX(F2:F6)');

    const minCell = sheet1.cells[9]?.find((c: any) => c.formula);
    expect(minCell?.formula).toBe('MIN(F2:F6)');
  });

  it('parses boolean values', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet3 = semantic.sheets[2];
    // 布尔值行 (row index 1)
    expect(sheet3.cells[1][1].type).toBe('boolean');
    expect(sheet3.cells[1][1].value).toBe(true);

    expect(sheet3.cells[1][2].type).toBe('boolean');
    expect(sheet3.cells[1][2].value).toBe(false);
  });

  it('parses hyperlinks', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet1 = semantic.sheets[0];
    expect(sheet1.hyperlinks).toBeDefined();
    expect(sheet1.hyperlinks!.length).toBeGreaterThanOrEqual(2);
    expect(sheet1.hyperlinks![0].ref).toBe('B2');
    expect(sheet1.hyperlinks![0].tooltip).toBe('查看笔记本电脑详情');
  });

  it('parses autoFilter', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet1 = semantic.sheets[0];
    expect(sheet1.autoFilter).toBeDefined();
    expect(sheet1.autoFilter!.ref).toBe('A1:G6');
  });

  it('parses dataValidations', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet3 = semantic.sheets[2];
    expect(sheet3.dataValidations).toBeDefined();
    expect(sheet3.dataValidations!.length).toBe(2);

    // 整数验证
    const wholeValidation = sheet3.dataValidations!.find((v: any) => v.type === 'whole');
    expect(wholeValidation).toBeDefined();
    expect(wholeValidation!.sqref).toBe('B9');

    // 列表验证
    const listValidation = sheet3.dataValidations!.find((v: any) => v.type === 'list');
    expect(listValidation).toBeDefined();
    expect(listValidation!.sqref).toBe('B10');
  });

  it('parses conditionalFormats', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet3 = semantic.sheets[2];
    expect(sheet3.conditionalFormats).toBeDefined();
    expect(sheet3.conditionalFormats!.length).toBeGreaterThanOrEqual(1);
    expect(sheet3.conditionalFormats![0].sqref).toBe('B13:E13');
  });

  it('parses printArea (pageMargins)', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const sheet3 = semantic.sheets[2];
    expect(sheet3.printArea).toBeDefined();
    expect(sheet3.printArea!.pageMargins).toBeDefined();
    expect(sheet3.printArea!.pageMargins!.left).toBe(0.7);
    expect(sheet3.printArea!.pageMargins!.right).toBe(0.7);
    expect(sheet3.printArea!.pageMargins!.top).toBe(0.75);
    expect(sheet3.printArea!.pageMargins!.bottom).toBe(0.75);
  });

  it('round-trips XLSX', async () => {
    const { semantic } = await loadXlsx(join(INPUT_DIR, 'test.xlsx'));

    const output = await serializeXlsx(semantic);
    expect(output).toBeInstanceOf(ArrayBuffer);

    const { semantic: restored } = await parseXlsx(output);
    expect(restored.sheets).toHaveLength(3);
    expect(restored.sheets[0].name).toBe('销售数据');
  });
});

// ============================================================
// PPTX 测试
// ============================================================
describe('Real PPTX file', () => {
  it('parses pptx meta correctly', async () => {
    const { raw, semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    expect(raw.parts.size).toBeGreaterThan(0);
    expect(semantic.meta.title).toBe('产品介绍演示文稿');
    expect(semantic.meta.creator).toBe('赵六');
  });

  it('parses 4 slides', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    expect(semantic.slides).toHaveLength(4);
  });

  it('parses slide 1 - title and subtitle', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const slide1 = semantic.slides[0];
    expect(slide1.elements).toHaveLength(2);

    expect(slide1.elements[0].type).toBe('text');
    const title = slide1.elements[0] as any;
    expect(title.paragraphs[0].runs[0].text).toBe('产品介绍演示文稿');

    expect(slide1.elements[1].type).toBe('text');
    const subtitle = slide1.elements[1] as any;
    expect(subtitle.paragraphs[0].runs[0].text).toBe('2024年度产品线全面介绍');
  });

  it('parses slide 2 - multiple product shapes', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const slide2 = semantic.slides[1];
    expect(slide2.elements).toHaveLength(4);

    const title = slide2.elements[0] as any;
    expect(title.paragraphs[0].runs[0].text).toBe('核心产品线');

    const product1 = slide2.elements[1] as any;
    expect(product1.paragraphs[0].runs[0].text).toBe('笔记本电脑系列');

    const product2 = slide2.elements[2] as any;
    expect(product2.paragraphs[0].runs[0].text).toBe('显示器系列');
  });

  it('parses slide 3 - content', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const slide3 = semantic.slides[2];
    expect(slide3.elements).toHaveLength(2);

    const title = slide3.elements[0] as any;
    expect(title.paragraphs[0].runs[0].text).toBe('总结与展望');

    const content = slide3.elements[1] as any;
    expect(content.paragraphs[0].runs[0].text).toBe('感谢关注！');
  });

  it('parses slide 4 - image, table, group, hyperlinks', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const slide4 = semantic.slides[3];
    expect(slide4.elements.length).toBeGreaterThanOrEqual(4);

    // 验证图片形状
    const imageShape = slide4.elements.find((e: any) => e.type === 'image');
    expect(imageShape).toBeDefined();
    expect((imageShape as any).relationshipId).toBe('rId1');

    // 验证表格形状
    const tableShape = slide4.elements.find((e: any) => e.type === 'table');
    expect(tableShape).toBeDefined();

    // 验证组合形状
    const groupShape = slide4.elements.find((e: any) => e.type === 'group');
    expect(groupShape).toBeDefined();
  });

  it('parses slide transitions', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const slide4 = semantic.slides[3];
    expect(slide4.transition).toBeDefined();
    expect(slide4.transition!.type).toBe('fade');
  });

  it('parses slide animations', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const slide4 = semantic.slides[3];
    expect(slide4.animations).toBeDefined();
    expect(slide4.animations!.length).toBeGreaterThanOrEqual(1);
  });

  it('parses slide notes', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const slide1 = semantic.slides[0];
    expect(slide1.notes).toBeDefined();
    expect(slide1.notes).toContain('备注');
  });

  it('round-trips PPTX', async () => {
    const { semantic } = await loadPptx(join(INPUT_DIR, 'test.pptx'));

    const output = await serializePptx(semantic);
    expect(output).toBeInstanceOf(ArrayBuffer);

    const { semantic: restored } = await parsePptx(output);
    expect(restored.slides).toHaveLength(4);
    expect(restored.slides[0].elements).toHaveLength(2);
  });
});
