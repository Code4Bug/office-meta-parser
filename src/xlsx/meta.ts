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
import type { XlsxWorkbook } from './types.js';

export const updateXlsxTitle = (wb: XlsxWorkbook, v: string) => _updateTitle(wb, v);
export const updateXlsxSubject = (wb: XlsxWorkbook, v: string) => _updateSubject(wb, v);
export const updateXlsxCreator = (wb: XlsxWorkbook, v: string) => _updateCreator(wb, v);
export const updateXlsxDescription = (wb: XlsxWorkbook, v: string) => _updateDescription(wb, v);
export const updateXlsxKeywords = (wb: XlsxWorkbook, v: string) => _updateKeywords(wb, v);
export const updateXlsxLastModifiedBy = (wb: XlsxWorkbook, v: string) => _updateLastModifiedBy(wb, v);
export const updateXlsxCategory = (wb: XlsxWorkbook, v: string) => _updateCategory(wb, v);

export const toXlsxJSON = (wb: XlsxWorkbook): XlsxWorkbook => _toJSON(wb);
export const toXlsxJSONString = (wb: XlsxWorkbook, space?: number): string => _toJSONString(wb, space);
export const saveXlsxJSON = (wb: XlsxWorkbook, path: string, space?: number): Promise<void> => _saveToJSON(wb, path, space);

export const xlsx = {
  updateTitle: updateXlsxTitle,
  updateSubject: updateXlsxSubject,
  updateCreator: updateXlsxCreator,
  updateDescription: updateXlsxDescription,
  updateKeywords: updateXlsxKeywords,
  updateLastModifiedBy: updateXlsxLastModifiedBy,
  updateCategory: updateXlsxCategory,
  toJSON: toXlsxJSON,
  toJSONString: toXlsxJSONString,
  saveJSON: saveXlsxJSON,
};
