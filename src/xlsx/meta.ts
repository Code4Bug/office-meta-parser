import { createMetaOps } from '../core/meta.js';
import type { XlsxWorkbook } from './types.js';

const ops = createMetaOps<XlsxWorkbook>();

export const updateXlsxTitle = ops.updateTitle;
export const updateXlsxSubject = ops.updateSubject;
export const updateXlsxCreator = ops.updateCreator;
export const updateXlsxDescription = ops.updateDescription;
export const updateXlsxKeywords = ops.updateKeywords;
export const updateXlsxLastModifiedBy = ops.updateLastModifiedBy;
export const updateXlsxCategory = ops.updateCategory;

export const toXlsxJSON = ops.toJSON;
export const toXlsxJSONString = ops.toJSONString;
export const saveXlsxJSON = ops.saveJSON;

export const xlsx = {
  updateTitle: ops.updateTitle,
  updateSubject: ops.updateSubject,
  updateCreator: ops.updateCreator,
  updateDescription: ops.updateDescription,
  updateKeywords: ops.updateKeywords,
  updateLastModifiedBy: ops.updateLastModifiedBy,
  updateCategory: ops.updateCategory,
  toJSON: ops.toJSON,
  toJSONString: ops.toJSONString,
  saveJSON: ops.saveJSON,
};
