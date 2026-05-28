import { createMetaOps } from '../core/meta.js';
import type { PptxPresentation } from './types.js';

const ops = createMetaOps<PptxPresentation>();

export const updatePptxTitle = ops.updateTitle;
export const updatePptxSubject = ops.updateSubject;
export const updatePptxCreator = ops.updateCreator;
export const updatePptxDescription = ops.updateDescription;
export const updatePptxKeywords = ops.updateKeywords;
export const updatePptxLastModifiedBy = ops.updateLastModifiedBy;
export const updatePptxCategory = ops.updateCategory;

export const toPptxJSON = ops.toJSON;
export const toPptxJSONString = ops.toJSONString;
export const savePptxJSON = ops.saveJSON;

export const pptx = {
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
