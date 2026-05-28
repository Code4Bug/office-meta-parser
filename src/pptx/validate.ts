import type { ValidationIssue } from '../core/validate.js';
import type { PptxPresentation, Slide, SlideElement, TextShape, GroupShape, TableShape } from './types.js';

export function validatePptx(pres: PptxPresentation): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  // ═══════════════ 必须项 (error) ═══════════════

  // 1. theme 不能缺失（Office 无法打开没有 theme 的 pptx）
  if (!pres.theme) {
    issues.push({ level: 'error', path: 'theme', message: 'theme is required (Office cannot open pptx without theme)' });
  }

  // 2. masters 不能为空
  if (!pres.masters || pres.masters.length === 0) {
    issues.push({ level: 'error', path: 'masters', message: 'at least one slideMaster is required' });
  }

  // 3. layouts 不能为空
  if (!pres.layouts || pres.layouts.length === 0) {
    issues.push({ level: 'error', path: 'layouts', message: 'at least one slideLayout is required' });
  }

  // 4. slides 不能为空
  if (!pres.slides || pres.slides.length === 0) {
    issues.push({ level: 'error', path: 'slides', message: 'at least one slide is required' });
  }

  // 5. 检查 shape 结构
  for (let i = 0; i < (pres.slides || []).length; i++) {
    validateSlide(pres.slides[i], `slides[${i}]`, issues);
  }

  // ═══════════════ 可选项 (warning) ═══════════════

  // 6. slide 引用的 layout 名称建议存在
  const layoutNames = new Set((pres.layouts || []).map(l => l.name).filter(Boolean));
  for (let i = 0; i < (pres.slides || []).length; i++) {
    const slide = pres.slides[i];
    if (slide.layout && !layoutNames.has(slide.layout)) {
      issues.push({ level: 'warning', path: `slides[${i}]`, message: `layout "${slide.layout}" not found in layouts` });
    }
  }

  // 7. slideSize 建议设置
  if (!pres.slideSize) {
    issues.push({ level: 'warning', path: 'slideSize', message: 'slideSize not set, will use default' });
  }

  return issues;
}

function validateSlide(slide: Slide, path: string, issues: ValidationIssue[]) {
  // elements 不能为空
  if (!slide.elements || slide.elements.length === 0) {
    issues.push({ level: 'warning', path, message: 'slide has no elements' });
  }
  for (let i = 0; i < (slide.elements || []).length; i++) {
    validateElement(slide.elements[i], `${path}.elements[${i}]`, issues);
  }
}

function validateElement(el: SlideElement, path: string, issues: ValidationIssue[]) {
  if (el.type === 'text') {
    validateTextShape(el, path, issues);
  } else if (el.type === 'group') {
    validateGroupShape(el, path, issues);
  } else if (el.type === 'table') {
    validateTableShape(el, path, issues);
  } else if (el.type === 'image') {
    // image 必须有 relationshipId
    if (!el.relationshipId) {
      issues.push({ level: 'error', path, message: 'image missing relationshipId' });
    }
  }
}

function validateTextShape(shape: TextShape, path: string, issues: ValidationIssue[]) {
  // txBody 的 paragraphs 不能为空（序列化器会补空 <a:p>，但应警告）
  if (!shape.paragraphs || shape.paragraphs.length === 0) {
    issues.push({ level: 'warning', path, message: 'text shape has empty paragraphs (serializer will add empty <a:p>)' });
  }
}

function validateGroupShape(shape: GroupShape, path: string, issues: ValidationIssue[]) {
  if (!shape.children || shape.children.length === 0) {
    issues.push({ level: 'warning', path, message: 'group shape has no children' });
  }
  for (let i = 0; i < (shape.children || []).length; i++) {
    validateElement(shape.children[i], `${path}.children[${i}]`, issues);
  }
}

function validateTableShape(shape: TableShape, path: string, issues: ValidationIssue[]) {
  // table rows 不能为空
  if (!shape.rows || shape.rows.length === 0) {
    issues.push({ level: 'error', path, message: 'table has no rows' });
  }
  for (let i = 0; i < (shape.rows || []).length; i++) {
    const row = shape.rows[i];
    if (!row.cells || row.cells.length === 0) {
      issues.push({ level: 'error', path: `${path}.rows[${i}]`, message: 'table row has no cells' });
    }
  }
}
