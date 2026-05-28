import { createMetaOps } from '../core/meta.js';
import type { DocxDocument } from './types.js';

const ops = createMetaOps<DocxDocument>();

export const updateDocxTitle = ops.updateTitle;
export const updateDocxSubject = ops.updateSubject;
export const updateDocxCreator = ops.updateCreator;
export const updateDocxDescription = ops.updateDescription;
export const updateDocxKeywords = ops.updateKeywords;
export const updateDocxLastModifiedBy = ops.updateLastModifiedBy;
export const updateDocxCategory = ops.updateCategory;

export const toDocxJSON = ops.toJSON;
export const toDocxJSONString = ops.toJSONString;
export const saveDocxJSON = ops.saveJSON;

export const docx = {
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
