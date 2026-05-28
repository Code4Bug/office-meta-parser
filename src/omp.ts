/**
 * OMP — 统一命名空间入口
 *
 * 通用 API：OMP.xxx()
 * 格式专属：OMP.docx.xxx() / OMP.xlsx.xxx() / OMP.pptx.xxx()
 */

// ---- Core ----
import { detectFormat, validate } from './core/validate.js';
import { loadFromFile, saveToFile, toBuffer, toJSON, toJSONString, saveToJSON } from './core/io.js';

// ---- DOCX ----
import { parseDocx, serializeDocx, loadDocx, saveDocx, createDocx, validateDocx } from './docx/index.js';
import {
  addComment as docxAddComment, removeComment as docxRemoveComment,
  listComments as docxListComments, getCommentText as docxGetCommentText,
  markCommentDone, markCommentUndone,
} from './docx/comments.js';
import {
  markInsert, markDelete, addFormatChange, clearRevision,
  listRevisions, hasPendingRevisions,
  acceptAllInserts, acceptAllDeletes, rejectAllInserts, rejectAllDeletes,
} from './docx/revisions.js';
import {
  updateDocxTitle, updateDocxSubject, updateDocxCreator,
  updateDocxDescription, updateDocxKeywords, updateDocxLastModifiedBy,
  updateDocxCategory, toDocxJSON, toDocxJSONString, saveDocxJSON,
} from './docx/meta.js';

// ---- XLSX ----
import { parseXlsx, serializeXlsx, loadXlsx, saveXlsx, createXlsx, validateXlsx } from './xlsx/index.js';
import {
  addComment as xlsxAddComment, removeComment as xlsxRemoveComment,
  listComments as xlsxListComments, listSheetComments,
  getCommentText as xlsxGetCommentText, updateComment as xlsxUpdateComment,
} from './xlsx/comments.js';
import {
  updateXlsxTitle, updateXlsxSubject, updateXlsxCreator,
  updateXlsxDescription, updateXlsxKeywords, updateXlsxLastModifiedBy,
  updateXlsxCategory, toXlsxJSON, toXlsxJSONString, saveXlsxJSON,
} from './xlsx/meta.js';

// ---- PPTX ----
import { parsePptx, serializePptx, loadPptx, savePptx, createPptx, validatePptx } from './pptx/index.js';
import {
  addComment as pptxAddComment, removeComment as pptxRemoveComment,
  listComments as pptxListComments, listSlideComments,
  getCommentText as pptxGetCommentText,
} from './pptx/comments.js';
import {
  updatePptxTitle, updatePptxSubject, updatePptxCreator,
  updatePptxDescription, updatePptxKeywords, updatePptxLastModifiedBy,
  updatePptxCategory, toPptxJSON, toPptxJSONString, savePptxJSON,
} from './pptx/meta.js';

export const OMP = {
  // ========== 通用 ==========
  detectFormat,
  validate,
  toBuffer,
  toJSON,
  toJSONString,
  saveToJSON,
  loadFromFile,
  saveToFile,

  // ========== DOCX ==========
  docx: {
    parse: parseDocx,
    serialize: serializeDocx,
    load: loadDocx,
    save: saveDocx,
    create: createDocx,
    validate: validateDocx,

    // 批注
    addComment: docxAddComment,
    removeComment: docxRemoveComment,
    listComments: docxListComments,
    getCommentText: docxGetCommentText,
    markCommentDone,
    markCommentUndone,

    // 修订
    markInsert,
    markDelete,
    addFormatChange,
    clearRevision,
    listRevisions,
    hasPendingRevisions,
    acceptAllInserts,
    acceptAllDeletes,
    rejectAllInserts,
    rejectAllDeletes,

    // 元数据
    updateTitle: updateDocxTitle,
    updateSubject: updateDocxSubject,
    updateCreator: updateDocxCreator,
    updateDescription: updateDocxDescription,
    updateKeywords: updateDocxKeywords,
    updateLastModifiedBy: updateDocxLastModifiedBy,
    updateCategory: updateDocxCategory,

    // JSON
    toJSON: toDocxJSON,
    toJSONString: toDocxJSONString,
    saveJSON: saveDocxJSON,
  },

  // ========== XLSX ==========
  xlsx: {
    parse: parseXlsx,
    serialize: serializeXlsx,
    load: loadXlsx,
    save: saveXlsx,
    create: createXlsx,
    validate: validateXlsx,

    // 批注
    addComment: xlsxAddComment,
    removeComment: xlsxRemoveComment,
    listComments: xlsxListComments,
    listSheetComments,
    getCommentText: xlsxGetCommentText,
    updateComment: xlsxUpdateComment,

    // 元数据
    updateTitle: updateXlsxTitle,
    updateSubject: updateXlsxSubject,
    updateCreator: updateXlsxCreator,
    updateDescription: updateXlsxDescription,
    updateKeywords: updateXlsxKeywords,
    updateLastModifiedBy: updateXlsxLastModifiedBy,
    updateCategory: updateXlsxCategory,

    // JSON
    toJSON: toXlsxJSON,
    toJSONString: toXlsxJSONString,
    saveJSON: saveXlsxJSON,
  },

  // ========== PPTX ==========
  pptx: {
    parse: parsePptx,
    serialize: serializePptx,
    load: loadPptx,
    save: savePptx,
    create: createPptx,
    validate: validatePptx,

    // 批注
    addComment: pptxAddComment,
    removeComment: pptxRemoveComment,
    listComments: pptxListComments,
    listSlideComments,
    getCommentText: pptxGetCommentText,

    // 元数据
    updateTitle: updatePptxTitle,
    updateSubject: updatePptxSubject,
    updateCreator: updatePptxCreator,
    updateDescription: updatePptxDescription,
    updateKeywords: updatePptxKeywords,
    updateLastModifiedBy: updatePptxLastModifiedBy,
    updateCategory: updatePptxCategory,

    // JSON
    toJSON: toPptxJSON,
    toJSONString: toPptxJSONString,
    saveJSON: savePptxJSON,
  },
};
