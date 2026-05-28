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
import type { DocxDocument } from './types.js';

export const updateDocxTitle = (doc: DocxDocument, v: string) => _updateTitle(doc, v);
export const updateDocxSubject = (doc: DocxDocument, v: string) => _updateSubject(doc, v);
export const updateDocxCreator = (doc: DocxDocument, v: string) => _updateCreator(doc, v);
export const updateDocxDescription = (doc: DocxDocument, v: string) => _updateDescription(doc, v);
export const updateDocxKeywords = (doc: DocxDocument, v: string) => _updateKeywords(doc, v);
export const updateDocxLastModifiedBy = (doc: DocxDocument, v: string) => _updateLastModifiedBy(doc, v);
export const updateDocxCategory = (doc: DocxDocument, v: string) => _updateCategory(doc, v);

export const toDocxJSON = (doc: DocxDocument): DocxDocument => _toJSON(doc);
export const toDocxJSONString = (doc: DocxDocument, space?: number): string => _toJSONString(doc, space);
export const saveDocxJSON = (doc: DocxDocument, path: string, space?: number): Promise<void> => _saveToJSON(doc, path, space);

export const docx = {
  updateTitle: updateDocxTitle,
  updateSubject: updateDocxSubject,
  updateCreator: updateDocxCreator,
  updateDescription: updateDocxDescription,
  updateKeywords: updateDocxKeywords,
  updateLastModifiedBy: updateDocxLastModifiedBy,
  updateCategory: updateDocxCategory,
  toJSON: toDocxJSON,
  toJSONString: toDocxJSONString,
  saveJSON: saveDocxJSON,
};
