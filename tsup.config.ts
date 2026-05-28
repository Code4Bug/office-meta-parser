import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    'index': 'src/index.ts',
    'core/index': 'src/core/index.ts',
    'docx/index': 'src/docx/index.ts',
    'pptx/index': 'src/pptx/index.ts',
    'xlsx/index': 'src/xlsx/index.ts',
  },
  format: ['esm'],
  dts: true,
  clean: true,
  splitting: true,
});
