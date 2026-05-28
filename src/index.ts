export * from './core/index.js';
export { OMP } from './omp.js';

// Format-specific modules can also be imported from subpaths:
// import { parseDocx, loadDocx, saveDocx, createDocx, validateDocx, docx, updateDocxTitle, ... } from 'office-meta-parser/docx'
// import { parseXlsx, loadXlsx, saveXlsx, createXlsx, validateXlsx, xlsx, updateXlsxTitle, ... } from 'office-meta-parser/xlsx'
// import { parsePptx, loadPptx, savePptx, createPptx, validatePptx, pptx, updatePptxTitle, ... } from 'office-meta-parser/pptx'
//
// Or use the unified namespace:
// import { OMP } from 'office-meta-parser'
// OMP.detectFormat(buffer)
// OMP.docx.create({ title: '...' })
// OMP.xlsx.addComment(wb, 0, 'A1', 'author', 'text')
