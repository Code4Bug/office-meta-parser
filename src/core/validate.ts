export interface ValidationIssue {
  level: 'error' | 'warning';
  path: string;
  message: string;
}

export type Validator<T> = (doc: T) => ValidationIssue[];

export type OfficeFormat = 'docx' | 'xlsx' | 'pptx';

export class ValidationError extends Error {
  readonly issues: ValidationIssue[];
  constructor(issues: ValidationIssue[]) {
    super(`Validation failed:\n${formatValidationReport(issues)}`);
    this.name = 'ValidationError';
    this.issues = issues;
  }
}

export class FormatError extends Error {
  readonly format?: string;
  constructor(message: string, format?: string) {
    super(message);
    this.name = 'FormatError';
    this.format = format;
  }
}

export function formatValidationReport(issues: ValidationIssue[]): string {
  if (issues.length === 0) return '';
  const lines = issues.map(i => `[${i.level.toUpperCase()}] ${i.path}: ${i.message}`);
  return lines.join('\n');
}

export function throwOnError(issues: ValidationIssue[]): void {
  const errors = issues.filter(i => i.level === 'error');
  if (errors.length > 0) {
    throw new ValidationError(errors);
  }
}

export function warnOnWarning(issues: ValidationIssue[]): void {
  const warnings = issues.filter(i => i.level === 'warning');
  if (warnings.length > 0) {
    console.warn(formatValidationReport(warnings));
  }
}

/**
 * 从 ZIP 内容自动检测 Office 文件格式
 */
export async function detectFormat(buffer: ArrayBuffer): Promise<OfficeFormat | null> {
  const { unzip } = await import('./zip.js');
  try {
    const entries = await unzip(buffer);
    const paths = new Set(entries.map(e => e.path));

    if (paths.has('word/document.xml')) return 'docx';
    if (paths.has('xl/workbook.xml')) return 'xlsx';
    if (paths.has('ppt/presentation.xml')) return 'pptx';
    return null;
  } catch {
    return null;
  }
}

export interface UnifiedValidationResult {
  format: OfficeFormat | null;
  issues: ValidationIssue[];
  valid: boolean;
}

/**
 * 统一校验入口：自动检测格式并执行对应的 validator
 */
export async function validate(buffer: ArrayBuffer): Promise<UnifiedValidationResult> {
  const format = await detectFormat(buffer);
  if (!format) {
    return {
      format: null,
      issues: [{ level: 'error', path: '', message: '无法识别的 Office 文件格式' }],
      valid: false,
    };
  }

  const issues: ValidationIssue[] = [];
  try {
    if (format === 'docx') {
      const { parseDocx, validateDocx } = await import('../docx/index.js');
      const { semantic } = await parseDocx(buffer);
      issues.push(...validateDocx(semantic));
    } else if (format === 'xlsx') {
      const { parseXlsx, validateXlsx } = await import('../xlsx/index.js');
      const { semantic } = await parseXlsx(buffer);
      issues.push(...validateXlsx(semantic));
    } else if (format === 'pptx') {
      const { parsePptx, validatePptx } = await import('../pptx/index.js');
      const { semantic } = await parsePptx(buffer);
      issues.push(...validatePptx(semantic));
    }
  } catch (error: any) {
    issues.push({ level: 'error', path: '', message: `解析失败: ${error.message}` });
  }

  return {
    format,
    issues,
    valid: issues.filter(i => i.level === 'error').length === 0,
  };
}
