# 测试文件说明

本目录包含用于测试 office-meta-parser 的测试文件和脚本。

## 测试文件

### 生成的测试文件

- `test.docx` - 综合 DOCX 测试文件
- `test.xlsx` - 综合 XLSX 测试文件
- `test.pptx` - 综合 PPTX 测试文件

这些文件由 `generate-fixtures.ts` 生成，覆盖了所有支持的属性点。

### 示例文件

- `demo-*.docx/xlsx/pptx` - 示例文档
- `判决书示例.docx` - 实际文档示例

## 测试脚本

### generate-fixtures.ts

生成综合测试文件的脚本，覆盖所有支持的 OOXML 特性：

- **DOCX**: 元数据、样式、段落、格式、表格（合并、垂直对齐）、页眉页脚、批注、超链接、编号列表
- **XLSX**: 元数据、多工作表、共享字符串、内联字符串、数字、公式、合并单元格、布尔值、超链接、自动筛选、数据验证、条件格式、打印区域
- **PPTX**: 元数据、多幻灯片、文本形状、图片形状、表格形状、组合形状、样式、过渡动画、入场动画、超链接、备注、幻灯片母版/布局、主题

运行方式：

```bash
npx tsx tests/files/generate-fixtures.ts
```

### codec-test.ts

编解码测试脚本，用于验证解析和序列化的正确性。

功能：
1. 加载测试文件（test.docx, test.xlsx, test.pptx）
2. 解析文件并提取语义数据
3. 将语义数据序列化回原格式
4. 重新解析验证数据完整性
5. 将所有过程数据保存到 `tests/logs/` 目录

运行方式：

```bash
# 使用 npm 脚本（推荐）
npm run test:codec

# 或直接运行
npx tsx tests/files/codec-test.ts
```

### real-files.test.ts

真实文件解析测试，验证所有支持的属性点都能正确解析。

运行方式：

```bash
npm test
```

## 日志输出

`codec-test.ts` 运行后会在 `tests/logs/` 目录下生成以下文件：

### 解析结果

- `docx-parse-result.json` - DOCX 解析摘要
- `xlsx-parse-result.json` - XLSX 解析摘要
- `pptx-parse-result.json` - PPTX 解析摘要

### 完整语义数据

- `docx-semantic-full.json` - DOCX 完整语义数据
- `xlsx-semantic-full.json` - XLSX 完整语义数据
- `pptx-semantic-full.json` - PPTX 完整语义数据

### 序列化结果

- `docx-serialized.docx` - 序列化后的 DOCX 文件
- `xlsx-serialized.xlsx` - 序列化后的 XLSX 文件
- `pptx-serialized.pptx` - 序列化后的 PPTX 文件

### 重新解析验证

- `docx-restored-semantic.json` - 从序列化文件重新解析的数据
- `xlsx-restored-semantic.json` - 从序列化文件重新解析的数据
- `pptx-restored-semantic.json` - 从序列化文件重新解析的数据

### 测试报告

- `test-summary.json` - 测试总结报告，包含：
  - 测试时间戳
  - 通过/失败数量
  - 每个格式的解析和序列化耗时
  - 错误信息（如有）

## 添加新的测试属性

1. 在 `src/*/types.ts` 中定义新的类型
2. 在 `src/*/parsers/` 中实现解析逻辑
3. 在 `src/*/serializers/` 中实现序列化逻辑
4. 在 `generate-fixtures.ts` 中添加测试数据
5. 运行 `npx tsx tests/files/generate-fixtures.ts` 重新生成测试文件
6. 在 `real-files.test.ts` 中添加测试用例
7. 运行 `npm run test:codec` 验证编解码正确性
8. 运行 `npm test` 确保所有测试通过
