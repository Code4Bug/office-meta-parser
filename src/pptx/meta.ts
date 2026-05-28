import {
  updateTitle as _updateTitle,
  updateSubject as _updateSubject,
  updateCreator as _updateCreator,
  updateDescription as _updateDescription,
  updateKeywords as _updateKeywords,
  updateLastModifiedBy as _updateLastModifiedBy,
  updateCategory as _updateCategory,
} from '../core/meta.js';
import { toJSON as _toJSON, toJSONString as _toJSONString, saveToJSON as _saveToJSON } from '../core/io.js';
import type { PptxPresentation } from './types.js';

export const updatePptxTitle = (pres: PptxPresentation, v: string) => _updateTitle(pres, v);
export const updatePptxSubject = (pres: PptxPresentation, v: string) => _updateSubject(pres, v);
export const updatePptxCreator = (pres: PptxPresentation, v: string) => _updateCreator(pres, v);
export const updatePptxDescription = (pres: PptxPresentation, v: string) => _updateDescription(pres, v);
export const updatePptxKeywords = (pres: PptxPresentation, v: string) => _updateKeywords(pres, v);
export const updatePptxLastModifiedBy = (pres: PptxPresentation, v: string) => _updateLastModifiedBy(pres, v);
export const updatePptxCategory = (pres: PptxPresentation, v: string) => _updateCategory(pres, v);

export const toPptxJSON = (pres: PptxPresentation): PptxPresentation => _toJSON(pres);
export const toPptxJSONString = (pres: PptxPresentation, space?: number): string => _toJSONString(pres, space);
export const savePptxJSON = (pres: PptxPresentation, path: string, space?: number): Promise<void> => _saveToJSON(pres, path, space);

export const pptx = {
  updateTitle: updatePptxTitle,
  updateSubject: updatePptxSubject,
  updateCreator: updatePptxCreator,
  updateDescription: updatePptxDescription,
  updateKeywords: updatePptxKeywords,
  updateLastModifiedBy: updatePptxLastModifiedBy,
  updateCategory: updatePptxCategory,
  toJSON: toPptxJSON,
  toJSONString: toPptxJSONString,
  saveJSON: savePptxJSON,
};
