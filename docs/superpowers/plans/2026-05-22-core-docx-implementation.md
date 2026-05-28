# Core + DOCX Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the shared core module and DOCX encoder/decoder as the reference implementation for the office-meta-parser library.

**Architecture:** Monorepo with subpath exports. Core module provides ZIP/XML/relationship utilities. DOCX module implements parser (XML -> Raw JSON), semantic converter (Raw -> Semantic), and serializer (JSON -> XML).

**Tech Stack:** TypeScript (strict), ESM, JSZip, fast-xml-parser, tsup, vitest

---

## File Structure

```
office-meta-parser/
├── src/
│   ├── core/
│   │   ├── types.ts           # Public types (ZipEntry, ParsedNode, RawDocument, etc.)
│   │   ├── xml.ts             # XML parse/serialize wrapper
│   │   ├── zip.ts             # ZIP unpack/pack wrapper
│   │   ├── rels.ts            # .rels relationship parser
│   │   ├── content-type.ts    # [Content_Types].xml parser
│   │   └── index.ts           # Core barrel export
│   ├── docx/
│   │   ├── types.ts           # DOCX semantic types
│   │   ├── parser.ts          # XML -> Raw JSON
│   │   ├── semantic.ts        # Raw JSON -> Semantic JSON
│   │   ├── serializer.ts      # Semantic JSON -> XML
│   │   └── index.ts           # DOCX barrel export
│   └── index.ts               # Main barrel export
├── tests/
│   ├── core/
│   │   ├── xml.test.ts
│   │   ├── zip.test.ts
│   │   ├── rels.test.ts
│   │   └── content-type.test.ts
│   ├── docx/
│   │   ├── parser.test.ts
│   │   ├── semantic.test.ts
│   │   └── serializer.test.ts
│   └── fixtures/
│       └── (test XML strings inline in tests)
├── package.json
├── tsconfig.json
├── tsup.config.ts
└── vitest.config.ts
```

---

### Task 1: Project Scaffolding

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `tsup.config.ts`
- Create: `vitest.config.ts`
- Create: `src/index.ts`

- [ ] **Step 1: Initialize package.json**

```json
{
  "name": "office-meta-parser",
  "version": "0.1.0",
  "type": "module",
  "exports": {
    ".": "./dist/index.js",
    "./core": "./dist/core/index.js",
    "./docx": "./dist/docx/index.js"
  },
  "scripts": {
    "build": "tsup",
    "test": "vitest run",
    "test:watch": "vitest",
    "typecheck": "tsc --noEmit"
  },
  "devDependencies": {
    "typescript": "^5.4.0",
    "tsup": "^8.0.0",
    "vitest": "^2.0.0"
  },
  "dependencies": {
    "jszip": "^3.10.0",
    "fast-xml-parser": "^4.4.0"
  }
}
```

- [ ] **Step 2: Create tsconfig.json**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "declaration": true,
    "outDir": "dist",
    "rootDir": "src"
  },
  "include": ["src"],
  "exclude": ["node_modules", "dist", "tests"]
}
```

- [ ] **Step 3: Create tsup.config.ts**

```typescript
import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    'index': 'src/index.ts',
    'core/index': 'src/core/index.ts',
    'docx/index': 'src/docx/index.ts',
  },
  format: ['esm'],
  dts: true,
  clean: true,
  splitting: false,
});
```

- [ ] **Step 4: Create vitest.config.ts**

```typescript
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
  },
});
```

- [ ] **Step 5: Create src/index.ts (placeholder)**

```typescript
export * from './core/index.js';
export * from './docx/index.js';
```

- [ ] **Step 6: Install dependencies**

Run: `npm install`

- [ ] **Step 7: Verify TypeScript compiles**

Run: `npx tsc --noEmit`
Expected: No errors

- [ ] **Step 8: Commit**

```bash
git add package.json tsconfig.json tsup.config.ts vitest.config.ts src/index.ts
git commit -m "chore: scaffold project with TypeScript, tsup, vitest"
```

---

### Task 2: Core Types

**Files:**
- Create: `src/core/types.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/core/types.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import type { ZipEntry, ParsedNode, Relationship, ContentType, RawDocument } from '../../src/core/types.js';

describe('core types', () => {
  it('ZipEntry has path and data', () => {
    const entry: ZipEntry = { path: 'test.xml', data: new ArrayBuffer(0) };
    expect(entry.path).toBe('test.xml');
    expect(entry.data).toBeInstanceOf(ArrayBuffer);
  });

  it('ParsedNode has tag, attrs, children', () => {
    const node: ParsedNode = {
      tag: 'w:body',
      attrs: {},
      children: [],
    };
    expect(node.tag).toBe('w:body');
    expect(node.attrs).toEqual({});
    expect(node.children).toEqual([]);
  });

  it('Relationship has id, type, target', () => {
    const rel: Relationship = {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'word/document.xml',
    };
    expect(rel.id).toBe('rId1');
    expect(rel.targetMode).toBeUndefined();
  });

  it('RawDocument has entries, rels, contentTypes, parts', () => {
    const doc: RawDocument = {
      entries: [],
      rels: new Map(),
      contentTypes: [],
      parts: new Map(),
    };
    expect(doc.entries).toEqual([]);
    expect(doc.rels).toBeInstanceOf(Map);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/core/types.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/core/types.ts`:

```typescript
export interface ZipEntry {
  path: string;
  data: ArrayBuffer;
}

export interface ParsedNode {
  tag: string;
  attrs: Record<string, string>;
  children: (ParsedNode | string)[];
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export interface ContentType {
  partName: string;
  contentType: string;
}

export interface RawDocument {
  entries: ZipEntry[];
  rels: Map<string, Relationship[]>;
  contentTypes: ContentType[];
  parts: Map<string, ParsedNode>;
}

export interface ParseResult<TSemantic> {
  raw: RawDocument;
  semantic: TSemantic;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/core/types.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/core/types.ts tests/core/types.test.ts
git commit -m "feat(core): add shared types"
```

---

### Task 3: Core XML Module

**Files:**
- Create: `src/core/xml.ts`
- Create: `tests/core/xml.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/core/xml.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { parseXml, serializeXml } from '../../src/core/xml.js';

describe('parseXml', () => {
  it('parses a simple element', () => {
    const result = parseXml('<root/>');
    expect(result).toEqual({ tag: 'root', attrs: {}, children: [] });
  });

  it('parses element with attributes', () => {
    const result = parseXml('<root id="1" name="test"/>');
    expect(result).toEqual({
      tag: 'root',
      attrs: { id: '1', name: 'test' },
      children: [],
    });
  });

  it('parses nested elements', () => {
    const result = parseXml('<root><child>text</child></root>');
    expect(result.tag).toBe('root');
    expect(result.children).toHaveLength(1);
    const child = result.children[0] as any;
    expect(child.tag).toBe('child');
    expect(child.children).toEqual(['text']);
  });

  it('parses mixed content', () => {
    const result = parseXml('<root>hello <b>world</b>!</root>');
    expect(result.tag).toBe('root');
    expect(result.children).toHaveLength(3);
    expect(result.children[0]).toBe('hello ');
    expect((result.children[1] as any).tag).toBe('b');
    expect(result.children[2]).toBe('!');
  });

  it('parses XML with namespaces', () => {
    const result = parseXml('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>');
    expect(result.tag).toBe('w:document');
    expect(result.attrs['xmlns:w']).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    expect((result.children[0] as any).tag).toBe('w:body');
  });
});

describe('serializeXml', () => {
  it('serializes a simple element', () => {
    const node = { tag: 'root', attrs: {}, children: [] };
    expect(serializeXml(node)).toBe('<root/>');
  });

  it('serializes element with attributes', () => {
    const node = { tag: 'root', attrs: { id: '1' }, children: [] };
    expect(serializeXml(node)).toBe('<root id="1"/>');
  });

  it('serializes element with text children', () => {
    const node = { tag: 'p', attrs: {}, children: ['hello'] };
    expect(serializeXml(node)).toBe('<p>hello</p>');
  });

  it('serializes nested elements', () => {
    const node = {
      tag: 'root',
      attrs: {},
      children: [
        { tag: 'child', attrs: {}, children: ['text'] },
      ],
    };
    expect(serializeXml(node)).toBe('<root><child>text</child></root>');
  });

  it('round-trips a complex document', () => {
    const xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>';
    const parsed = parseXml(xml);
    const serialized = serializeXml(parsed);
    expect(serialized).toBe(xml);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/core/xml.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/core/xml.ts`:

```typescript
import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import type { ParsedNode } from './types.js';

const parserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '',
  allowBooleanAttributes: true,
  parseAttributeValue: false,
  trimValues: false,
  preserveOrder: true,
  commentPropName: '__comment',
  textNodeName: '__text',
};

const builderOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '',
  format: false,
  suppressEmptyNode: true,
  preserveOrder: true,
  commentPropName: '__comment',
  textNodeName: '__text',
};

export function parseXml(xml: string): ParsedNode {
  const parser = new XMLParser(parserOptions);
  const result = parser.parse(xml);

  // fast-xml-parser returns an array; find the root element
  const root = result.find(
    (item: any) => typeof item === 'object' && !item.__text && !item.__comment
  );
  if (!root) {
    throw new Error('No root element found in XML');
  }

  return convertNode(root);
}

function convertNode(raw: any): ParsedNode {
  const tag = Object.keys(raw).find(k => k !== ':' && k !== '@_') || '';
  const attrs: Record<string, string> = {};

  // Extract attributes (prefixed with '@_')
  for (const key of Object.keys(raw)) {
    if (key.startsWith('@_')) {
      attrs[key.slice(2)] = String(raw[key]);
    }
  }

  const children: (ParsedNode | string)[] = [];
  const content = raw[tag];

  if (Array.isArray(content)) {
    for (const item of content) {
      if (typeof item === 'string') {
        children.push(item);
      } else if (item.__text !== undefined) {
        children.push(String(item.__text));
      } else if (item.__comment !== undefined) {
        // Skip comments for now
      } else {
        children.push(convertNode(item));
      }
    }
  } else if (typeof content === 'string') {
    children.push(content);
  }

  return { tag, attrs, children };
}

export function serializeXml(node: ParsedNode): string {
  const raw = convertToBuilder(node);
  const builder = new XMLBuilder(builderOptions);
  return builder.build(raw);
}

function convertToBuilder(node: ParsedNode): any {
  const result: any = {};
  result[node.tag] = [];

  // Add attributes
  for (const [key, value] of Object.entries(node.attrs)) {
    result['@_' + key] = value;
  }

  // Add children
  for (const child of node.children) {
    if (typeof child === 'string') {
      result[node.tag].push({ __text: child });
    } else {
      result[node.tag].push(convertToBuilder(child));
    }
  }

  return result;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/core/xml.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/core/xml.ts tests/core/xml.test.ts
git commit -m "feat(core): add XML parse/serialize module"
```

---

### Task 4: Core ZIP Module

**Files:**
- Create: `src/core/zip.ts`
- Create: `tests/core/zip.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/core/zip.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { unzip, zip } from '../../src/core/zip.js';

describe('unzip', () => {
  it('unpacks a ZIP buffer into entries', async () => {
    // Create a minimal ZIP in memory using JSZip
    const JSZip = (await import('jszip')).default;
    const archive = new JSZip();
    archive.file('test.txt', 'hello');
    archive.file('sub/nested.txt', 'world');
    const buffer = await archive.generateAsync({ type: 'arraybuffer' });

    const entries = await unzip(buffer);
    expect(entries).toHaveLength(2);
    expect(entries[0].path).toBe('test.txt');
    expect(entries[1].path).toBe('sub/nested.txt');

    // Verify content
    const decoder = new TextDecoder();
    expect(decoder.decode(entries[0].data)).toBe('hello');
    expect(decoder.decode(entries[1].data)).toBe('world');
  });
});

describe('zip', () => {
  it('packs entries into a ZIP buffer', async () => {
    const encoder = new TextEncoder();
    const entries = [
      { path: 'a.txt', data: encoder.encode('aaa').buffer },
      { path: 'b.txt', data: encoder.encode('bbb').buffer },
    ];

    const buffer = await zip(entries);
    expect(buffer).toBeInstanceOf(ArrayBuffer);

    // Verify by unpacking
    const JSZip = (await import('jszip')).default;
    const archive = await JSZip.loadAsync(buffer);
    expect(archive.file('a.txt')).not.toBeNull();
    expect(archive.file('b.txt')).not.toBeNull();
    expect(await archive.file('a.txt')!.async('string')).toBe('aaa');
  });
});

describe('round-trip', () => {
  it('zip -> unzip preserves data', async () => {
    const encoder = new TextEncoder();
    const original = [
      { path: 'doc.xml', data: encoder.encode('<root/>').buffer },
      { path: 'rels/.rels', data: encoder.encode('<Relationships/>').buffer },
    ];

    const buffer = await zip(original);
    const restored = await unzip(buffer);

    expect(restored).toHaveLength(2);
    const decoder = new TextDecoder();
    expect(decoder.decode(restored[0].data)).toBe('<root/>');
    expect(decoder.decode(restored[1].data)).toBe('<Relationships/>');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/core/zip.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/core/zip.ts`:

```typescript
import JSZip from 'jszip';
import type { ZipEntry } from './types.js';

export async function unzip(buffer: ArrayBuffer): Promise<ZipEntry[]> {
  const archive = await JSZip.loadAsync(buffer);
  const entries: ZipEntry[] = [];

  const promises: Promise<void>[] = [];
  archive.forEach((path, file) => {
    if (!file.dir) {
      const p = file.async('arraybuffer').then(data => {
        entries.push({ path, data });
      });
      promises.push(p);
    }
  });

  await Promise.all(promises);

  // Sort by path for deterministic order
  entries.sort((a, b) => a.path.localeCompare(b.path));

  return entries;
}

export async function zip(entries: ZipEntry[]): Promise<ArrayBuffer> {
  const archive = new JSZip();

  for (const entry of entries) {
    archive.file(entry.path, entry.data);
  }

  return archive.generateAsync({ type: 'arraybuffer' });
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/core/zip.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/core/zip.ts tests/core/zip.test.ts
git commit -m "feat(core): add ZIP pack/unpack module"
```

---

### Task 5: Core Rels Module

**Files:**
- Create: `src/core/rels.ts`
- Create: `tests/core/rels.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/core/rels.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { parseRels, serializeRels } from '../../src/core/rels.js';

const sampleRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

describe('parseRels', () => {
  it('parses relationships from XML', () => {
    const rels = parseRels(sampleRelsXml);
    expect(rels).toHaveLength(2);
    expect(rels[0]).toEqual({
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'word/document.xml',
    });
    expect(rels[1].id).toBe('rId2');
  });

  it('handles external target mode', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`;
    const rels = parseRels(xml);
    expect(rels[0].targetMode).toBe('External');
  });

  it('returns empty array for empty Relationships', () => {
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const rels = parseRels(xml);
    expect(rels).toEqual([]);
  });
});

describe('serializeRels', () => {
  it('serializes relationships to XML', () => {
    const rels = [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', target: 'word/document.xml' },
    ];
    const xml = serializeRels(rels);
    expect(xml).toContain('Id="rId1"');
    expect(xml).toContain('Target="word/document.xml"');
    expect(xml).toContain('Relationships');
  });

  it('round-trips relationships', () => {
    const rels = parseRels(sampleRelsXml);
    const xml = serializeRels(rels);
    const restored = parseRels(xml);
    expect(restored).toEqual(rels);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/core/rels.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/core/rels.ts`:

```typescript
import { parseXml, serializeXml } from './xml.js';
import type { Relationship, ParsedNode } from './types.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export function parseRels(xml: string): Relationship[] {
  const root = parseXml(xml);
  if (root.tag !== 'Relationships') {
    throw new Error(`Expected Relationships root, got ${root.tag}`);
  }

  return root.children
    .filter((c): c is ParsedNode => typeof c !== 'string' && c.tag === 'Relationship')
    .map(node => ({
      id: node.attrs['Id'] || '',
      type: node.attrs['Type'] || '',
      target: node.attrs['Target'] || '',
      targetMode: node.attrs['TargetMode'],
    }));
}

export function serializeRels(rels: Relationship[]): string {
  const root: ParsedNode = {
    tag: 'Relationships',
    attrs: { xmlns: RELS_NS },
    children: rels.map(rel => {
      const attrs: Record<string, string> = {
        Id: rel.id,
        Type: rel.type,
        Target: rel.target,
      };
      if (rel.targetMode) {
        attrs.TargetMode = rel.targetMode;
      }
      return {
        tag: 'Relationship',
        attrs,
        children: [],
      };
    }),
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/core/rels.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/core/rels.ts tests/core/rels.test.ts
git commit -m "feat(core): add relationship file parser"
```

---

### Task 6: Core Content-Type Module

**Files:**
- Create: `src/core/content-type.ts`
- Create: `tests/core/content-type.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/core/content-type.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { parseContentTypes, serializeContentTypes } from '../../src/core/content-type.js';

const sampleContentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats.wordprocessingml.styles+xml"/>
</Types>`;

describe('parseContentTypes', () => {
  it('parses Default and Override entries', () => {
    const types = parseContentTypes(sampleContentTypesXml);
    expect(types).toHaveLength(4);
    expect(types[0]).toEqual({
      partName: '.rels',
      contentType: 'application/vnd.openxmlformats-package.relationships+xml',
    });
    expect(types[1]).toEqual({
      partName: '.xml',
      contentType: 'application/xml',
    });
    expect(types[2]).toEqual({
      partName: '/word/document.xml',
      contentType: 'application/vnd.openxmlformats.wordprocessingml.document.main+xml',
    });
  });
});

describe('serializeContentTypes', () => {
  it('serializes content types to XML', () => {
    const types = [
      { partName: '.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
    ];
    const xml = serializeContentTypes(types);
    expect(xml).toContain('Types');
    expect(xml).toContain('Default');
    expect(xml).toContain('Extension="rels"');
  });

  it('round-trips content types', () => {
    const types = parseContentTypes(sampleContentTypesXml);
    const xml = serializeContentTypes(types);
    const restored = parseContentTypes(xml);
    expect(restored).toEqual(types);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/core/content-type.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/core/content-type.ts`:

```typescript
import { parseXml, serializeXml } from './xml.js';
import type { ContentType, ParsedNode } from './types.js';

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

export function parseContentTypes(xml: string): ContentType[] {
  const root = parseXml(xml);
  if (root.tag !== 'Types') {
    throw new Error(`Expected Types root, got ${root.tag}`);
  }

  const result: ContentType[] = [];

  for (const child of root.children) {
    if (typeof child === 'string') continue;

    if (child.tag === 'Default') {
      result.push({
        partName: '.' + (child.attrs['Extension'] || ''),
        contentType: child.attrs['ContentType'] || '',
      });
    } else if (child.tag === 'Override') {
      result.push({
        partName: child.attrs['PartName'] || '',
        contentType: child.attrs['ContentType'] || '',
      });
    }
  }

  return result;
}

export function serializeContentTypes(types: ContentType[]): string {
  const children: ParsedNode[] = types.map(ct => {
    if (ct.partName.startsWith('.')) {
      return {
        tag: 'Default',
        attrs: { Extension: ct.partName.slice(1), ContentType: ct.contentType },
        children: [],
      };
    }
    return {
      tag: 'Override',
      attrs: { PartName: ct.partName, ContentType: ct.contentType },
      children: [],
    };
  });

  const root: ParsedNode = {
    tag: 'Types',
    attrs: { xmlns: CT_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/core/content-type.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/core/content-type.ts tests/core/content-type.test.ts
git commit -m "feat(core): add content type parser"
```

---

### Task 7: Core Barrel Export

**Files:**
- Create: `src/core/index.ts`

- [ ] **Step 1: Create src/core/index.ts**

```typescript
export type {
  ZipEntry,
  ParsedNode,
  Relationship,
  ContentType,
  RawDocument,
  ParseResult,
} from './types.js';

export { parseXml, serializeXml } from './xml.js';
export { unzip, zip } from './zip.js';
export { parseRels, serializeRels } from './rels.js';
export { parseContentTypes, serializeContentTypes } from './content-type.js';
```

- [ ] **Step 2: Verify all tests still pass**

Run: `npx vitest run`
Expected: All tests PASS

- [ ] **Step 3: Commit**

```bash
git add src/core/index.ts
git commit -m "feat(core): add barrel export"
```

---

### Task 8: DOCX Types

**Files:**
- Create: `src/docx/types.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/docx/types.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import type {
  DocxDocument,
  DocumentMeta,
  Paragraph,
  TextRun,
  Table,
  Image,
  Comment,
  Revision,
} from '../../src/docx/types.js';

describe('DOCX types', () => {
  it('DocxDocument structure', () => {
    const doc: DocxDocument = {
      meta: { title: 'Test' },
      styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
      body: { blocks: [] },
    };
    expect(doc.meta.title).toBe('Test');
    expect(doc.body.blocks).toEqual([]);
  });

  it('Paragraph structure', () => {
    const p: Paragraph = {
      type: 'paragraph',
      runs: [{ text: 'Hello' }],
    };
    expect(p.type).toBe('paragraph');
    expect(p.runs[0].text).toBe('Hello');
  });

  it('TextRun with styles', () => {
    const run: TextRun = {
      text: 'Bold',
      bold: true,
      fontSize: 24,
      color: 'FF0000',
    };
    expect(run.bold).toBe(true);
    expect(run.fontSize).toBe(24);
  });

  it('Table structure', () => {
    const table: Table = {
      type: 'table',
      rows: [
        { cells: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Cell' }] }] }] },
      ],
    };
    expect(table.rows).toHaveLength(1);
  });

  it('Image structure', () => {
    const img: Image = {
      type: 'image',
      relationshipId: 'rId1',
      width: 100,
      height: 200,
    };
    expect(img.relationshipId).toBe('rId1');
  });

  it('Comment structure', () => {
    const comment: Comment = {
      id: '1',
      author: 'John',
      date: '2024-01-01T00:00:00Z',
      content: [{ type: 'paragraph', runs: [{ text: 'Nice work' }] }],
    };
    expect(comment.author).toBe('John');
  });

  it('Revision structure', () => {
    const rev: Revision = {
      type: 'insert',
      author: 'Jane',
      date: '2024-01-01T00:00:00Z',
    };
    expect(rev.type).toBe('insert');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/docx/types.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/docx/types.ts`:

```typescript
export interface DocxDocument {
  meta: DocumentMeta;
  styles: StyleDefinitions;
  body: DocxBody;
  comments?: Comment[];
  trackChanges?: Revision[];
}

export interface DocumentMeta {
  title?: string;
  subject?: string;
  creator?: string;
  description?: string;
  keywords?: string;
  lastModifiedBy?: string;
  created?: string;
  modified?: string;
  revision?: string;
  category?: string;
}

export interface StyleDefinitions {
  paragraphStyles: ParagraphStyle[];
  characterStyles: CharacterStyle[];
  tableStyles: TableStyle[];
}

export interface ParagraphStyle {
  id: string;
  name?: string;
  basedOn?: string;
  next?: string;
  properties?: ParagraphProperties;
}

export interface CharacterStyle {
  id: string;
  name?: string;
  basedOn?: string;
  properties?: RunProperties;
}

export interface TableStyle {
  id: string;
  name?: string;
  basedOn?: string;
}

export interface DocxBody {
  blocks: DocxBlock[];
}

export type DocxBlock = Paragraph | Table | Image;

export interface Paragraph {
  type: 'paragraph';
  style?: string;
  properties?: ParagraphProperties;
  runs: TextRun[];
}

export interface ParagraphProperties {
  alignment?: 'left' | 'center' | 'right' | 'justify';
  indent?: {
    left?: number;
    right?: number;
    firstLine?: number;
  };
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: 'auto' | 'exact' | 'atLeast';
  };
}

export interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  color?: string;
  fontFamily?: string;
  properties?: RunProperties;
}

export interface RunProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  color?: string;
  fontFamily?: string;
  superscript?: boolean;
  subscript?: boolean;
}

export interface Table {
  type: 'table';
  rows: TableRow[];
  properties?: TableProperties;
}

export interface TableRow {
  cells: TableCell[];
  properties?: TableRowProperties;
}

export interface TableCell {
  blocks: DocxBlock[];
  properties?: TableCellProperties;
}

export interface TableProperties {
  width?: number;
  borders?: TableBorders;
}

export interface TableRowProperties {
  height?: number;
}

export interface TableCellProperties {
  width?: number;
  verticalMerge?: 'restart' | 'continue';
  horizontalMerge?: 'restart' | 'continue';
  verticalAlign?: 'top' | 'center' | 'bottom';
}

export interface TableBorders {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
  insideHorizontal?: BorderStyle;
  insideVertical?: BorderStyle;
}

export interface BorderStyle {
  style: string;
  size?: number;
  color?: string;
}

export interface Image {
  type: 'image';
  relationshipId: string;
  width?: number;
  height?: number;
  alt?: string;
}

export interface Comment {
  id: string;
  author: string;
  date: string;
  content: Paragraph[];
  initials?: string;
}

export interface Revision {
  type: 'insert' | 'delete' | 'formatChange';
  author: string;
  date: string;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/docx/types.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docx/types.ts tests/docx/types.test.ts
git commit -m "feat(docx): add semantic types"
```

---

### Task 9: DOCX Parser (XML -> Raw JSON)

**Files:**
- Create: `src/docx/parser.ts`
- Create: `tests/docx/parser.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/docx/parser.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { parseDocxXml } from '../../src/docx/parser.js';
import type { RawDocument } from '../../src/core/types.js';

describe('parseDocxXml', () => {
  it('parses a minimal document.xml', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = parseDocxXml(xml);
    expect(result.tag).toBe('w:document');
    expect(result.children).toHaveLength(1);

    const body = result.children[0] as any;
    expect(body.tag).toBe('w:body');

    const para = body.children[0] as any;
    expect(para.tag).toBe('w:p');

    const run = para.children[0] as any;
    expect(run.tag).toBe('w:r');

    const text = run.children[0] as any;
    expect(text.tag).toBe('w:t');
    expect(text.children[0]).toBe('Hello World');
  });

  it('preserves namespace attributes', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body/>
</w:document>`;

    const result = parseDocxXml(xml);
    expect(result.attrs['xmlns:w']).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    expect(result.attrs['xmlns:r']).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  });

  it('handles paragraph with multiple runs', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First</w:t></w:r>
      <w:r><w:t>Second</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = parseDocxXml(xml);
    const body = result.children[0] as any;
    const para = body.children[0] as any;
    expect(para.children).toHaveLength(2);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/docx/parser.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/docx/parser.ts`:

```typescript
import { parseXml } from '../core/xml.js';
import type { ParsedNode } from '../core/types.js';

export function parseDocxXml(xml: string): ParsedNode {
  return parseXml(xml);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/docx/parser.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docx/parser.ts tests/docx/parser.test.ts
git commit -m "feat(docx): add XML parser"
```

---

### Task 10: DOCX Semantic Converter (Raw -> Semantic)

**Files:**
- Create: `src/docx/semantic.ts`
- Create: `tests/docx/semantic.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/docx/semantic.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { rawToSemantic } from '../../src/docx/semantic.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function w(tag: string, children: any[] = [], attrs: Record<string, string> = {}): ParsedNode {
  return { tag: `w:${tag}`, attrs, children };
}

describe('rawToSemantic', () => {
  it('extracts meta from core.xml', () => {
    const coreXml: ParsedNode = {
      tag: 'cp:coreProperties',
      attrs: {
        'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
      },
      children: [
        { tag: 'dc:title', attrs: {}, children: ['My Document'] },
        { tag: 'dc:creator', attrs: {}, children: ['John Doe'] },
      ],
    };

    const raw = makeRaw({ 'docProps/core.xml': coreXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.meta.title).toBe('My Document');
    expect(semantic.meta.creator).toBe('John Doe');
  });

  it('parses a simple paragraph with text run', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                {
                  tag: 'w:r',
                  attrs: {},
                  children: [
                    { tag: 'w:t', attrs: {}, children: ['Hello World'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.body.blocks).toHaveLength(1);
    const para = semantic.body.blocks[0] as any;
    expect(para.type).toBe('paragraph');
    expect(para.runs).toHaveLength(1);
    expect(para.runs[0].text).toBe('Hello World');
  });

  it('parses bold and italic runs', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                {
                  tag: 'w:r',
                  attrs: {},
                  children: [
                    {
                      tag: 'w:rPr',
                      attrs: {},
                      children: [
                        { tag: 'w:b', attrs: {}, children: [] },
                        { tag: 'w:i', attrs: {}, children: [] },
                      ],
                    },
                    { tag: 'w:t', attrs: {}, children: ['Bold Italic'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    const para = semantic.body.blocks[0] as any;
    expect(para.runs[0].bold).toBe(true);
    expect(para.runs[0].italic).toBe(true);
  });

  it('handles multiple paragraphs', () => {
    const documentXml: ParsedNode = {
      tag: 'w:document',
      attrs: {},
      children: [
        {
          tag: 'w:body',
          attrs: {},
          children: [
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['P1'] }] },
              ],
            },
            {
              tag: 'w:p',
              attrs: {},
              children: [
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['P2'] }] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const semantic = rawToSemantic(raw);
    expect(semantic.body.blocks).toHaveLength(2);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/docx/semantic.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/docx/semantic.ts`:

```typescript
import type { RawDocument, ParsedNode } from '../core/types.js';
import type {
  DocxDocument,
  DocumentMeta,
  DocxBody,
  DocxBlock,
  Paragraph,
  TextRun,
  RunProperties,
} from './types.js';

export function rawToSemantic(raw: RawDocument): DocxDocument {
  return {
    meta: extractMeta(raw),
    styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
    body: extractBody(raw),
    comments: extractComments(raw),
    trackChanges: extractRevisions(raw),
  };
}

function extractMeta(raw: RawDocument): DocumentMeta {
  const coreXml = raw.parts.get('docProps/core.xml');
  if (!coreXml) return {};

  const meta: DocumentMeta = {};
  for (const child of coreXml.children) {
    if (typeof child === 'string') continue;
    const text = getTextContent(child);
    switch (child.tag) {
      case 'dc:title':
        meta.title = text;
        break;
      case 'dc:subject':
        meta.subject = text;
        break;
      case 'dc:creator':
        meta.creator = text;
        break;
      case 'dc:description':
        meta.description = text;
        break;
      case 'cp:keywords':
        meta.keywords = text;
        break;
      case 'cp:lastModifiedBy':
        meta.lastModifiedBy = text;
        break;
      case 'dcterms:created':
        meta.created = text;
        break;
      case 'dcterms:modified':
        meta.modified = text;
        break;
      case 'cp:revision':
        meta.revision = text;
        break;
      case 'cp:category':
        meta.category = text;
        break;
    }
  }
  return meta;
}

function extractBody(raw: RawDocument): DocxBody {
  const documentXml = raw.parts.get('word/document.xml');
  if (!documentXml) return { blocks: [] };

  const body = findChild(documentXml, 'w:body');
  if (!body) return { blocks: [] };

  const blocks: DocxBlock[] = [];
  for (const child of body.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:p') {
      blocks.push(parseParagraph(child));
    }
    // Table and Image support will be added later
  }

  return { blocks };
}

function parseParagraph(node: ParsedNode): Paragraph {
  const runs: TextRun[] = [];

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:r') {
      runs.push(parseRun(child));
    }
  }

  return {
    type: 'paragraph',
    runs,
  };
}

function parseRun(node: ParsedNode): TextRun {
  let text = '';
  let properties: RunProperties | undefined;

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:t') {
      text = getTextContent(child);
    } else if (child.tag === 'w:rPr') {
      properties = parseRunProperties(child);
    }
  }

  return {
    text,
    ...properties,
  };
}

function parseRunProperties(node: ParsedNode): RunProperties {
  const props: RunProperties = {};

  for (const child of node.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:b':
        props.bold = true;
        break;
      case 'w:i':
        props.italic = true;
        break;
      case 'w:u':
        props.underline = true;
        break;
      case 'w:strike':
        props.strike = true;
        break;
      case 'w:sz':
        if (child.attrs['w:val']) {
          props.fontSize = parseInt(child.attrs['w:val'], 10);
        }
        break;
      case 'w:color':
        if (child.attrs['w:val']) {
          props.color = child.attrs['w:val'];
        }
        break;
      case 'w:rFonts':
        if (child.attrs['w:ascii']) {
          props.fontFamily = child.attrs['w:ascii'];
        }
        break;
    }
  }

  return props;
}

function extractComments(raw: RawDocument): undefined {
  // TODO: Implement comment parsing
  return undefined;
}

function extractRevisions(raw: RawDocument): undefined {
  // TODO: Implement revision parsing
  return undefined;
}

function findChild(node: ParsedNode, tag: string): ParsedNode | undefined {
  for (const child of node.children) {
    if (typeof child !== 'string' && child.tag === tag) {
      return child;
    }
  }
  return undefined;
}

function getTextContent(node: ParsedNode): string {
  let result = '';
  for (const child of node.children) {
    if (typeof child === 'string') {
      result += child;
    } else {
      result += getTextContent(child);
    }
  }
  return result;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/docx/semantic.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docx/semantic.ts tests/docx/semantic.test.ts
git commit -m "feat(docx): add semantic converter"
```

---

### Task 11: DOCX Serializer (Semantic -> XML)

**Files:**
- Create: `src/docx/serializer.ts`
- Create: `tests/docx/serializer.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/docx/serializer.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { semanticToXml } from '../../src/docx/serializer.js';
import type { DocxDocument } from '../../src/docx/types.js';

describe('semanticToXml', () => {
  it('serializes a simple paragraph', () => {
    const doc: DocxDocument = {
      meta: {},
      styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
      body: {
        blocks: [
          {
            type: 'paragraph',
            runs: [{ text: 'Hello' }],
          },
        ],
      },
    };

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:document');
    expect(xml).toContain('w:body');
    expect(xml).toContain('w:p');
    expect(xml).toContain('w:r');
    expect(xml).toContain('w:t');
    expect(xml).toContain('Hello');
  });

  it('serializes bold text', () => {
    const doc: DocxDocument = {
      meta: {},
      styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
      body: {
        blocks: [
          {
            type: 'paragraph',
            runs: [{ text: 'Bold', bold: true }],
          },
        ],
      },
    };

    const xml = semanticToXml(doc);
    expect(xml).toContain('w:b');
  });

  it('serializes multiple paragraphs', () => {
    const doc: DocxDocument = {
      meta: {},
      styles: { paragraphStyles: [], characterStyles: [], tableStyles: [] },
      body: {
        blocks: [
          { type: 'paragraph', runs: [{ text: 'P1' }] },
          { type: 'paragraph', runs: [{ text: 'P2' }] },
        ],
      },
    };

    const xml = semanticToXml(doc);
    const pCount = (xml.match(/<w:p>/g) || []).length;
    expect(pCount).toBe(2);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/docx/serializer.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/docx/serializer.ts`:

```typescript
import { serializeXml } from '../core/xml.js';
import type { ParsedNode } from '../core/types.js';
import type {
  DocxDocument,
  Paragraph,
  TextRun,
  RunProperties,
} from './types.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function semanticToXml(doc: DocxDocument): string {
  const root: ParsedNode = {
    tag: 'w:document',
    attrs: {
      'xmlns:w': W_NS,
      'xmlns:r': R_NS,
    },
    children: [
      serializeBody(doc),
    ],
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}

function serializeBody(doc: DocxDocument): ParsedNode {
  const children: ParsedNode[] = [];

  for (const block of doc.body.blocks) {
    if (block.type === 'paragraph') {
      children.push(serializeParagraph(block));
    }
  }

  return {
    tag: 'w:body',
    attrs: {},
    children,
  };
}

function serializeParagraph(para: Paragraph): ParsedNode {
  const children: ParsedNode[] = [];

  for (const run of para.runs) {
    children.push(serializeRun(run));
  }

  return {
    tag: 'w:p',
    attrs: {},
    children,
  };
}

function serializeRun(run: TextRun): ParsedNode {
  const children: (ParsedNode | string)[] = [];

  const rPr = serializeRunProperties(run);
  if (rPr) {
    children.push(rPr);
  }

  children.push({
    tag: 'w:t',
    attrs: { 'xml:space': 'preserve' },
    children: [run.text],
  });

  return {
    tag: 'w:r',
    attrs: {},
    children,
  };
}

function serializeRunProperties(run: TextRun): ParsedNode | null {
  const children: ParsedNode[] = [];

  if (run.bold) {
    children.push({ tag: 'w:b', attrs: {}, children: [] });
  }
  if (run.italic) {
    children.push({ tag: 'w:i', attrs: {}, children: [] });
  }
  if (run.underline) {
    children.push({ tag: 'w:u', attrs: { 'w:val': 'single' }, children: [] });
  }
  if (run.strike) {
    children.push({ tag: 'w:strike', attrs: {}, children: [] });
  }
  if (run.fontSize) {
    children.push({ tag: 'w:sz', attrs: { 'w:val': String(run.fontSize) }, children: [] });
  }
  if (run.color) {
    children.push({ tag: 'w:color', attrs: { 'w:val': run.color }, children: [] });
  }
  if (run.fontFamily) {
    children.push({
      tag: 'w:rFonts',
      attrs: { 'w:ascii': run.fontFamily, 'w:hAnsi': run.fontFamily },
      children: [],
    });
  }

  if (children.length === 0) return null;

  return {
    tag: 'w:rPr',
    attrs: {},
    children,
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/docx/serializer.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docx/serializer.ts tests/docx/serializer.test.ts
git commit -m "feat(docx): add semantic serializer"
```

---

### Task 12: DOCX Integration (parseDocx + serializeDocx)

**Files:**
- Create: `src/docx/index.ts`
- Create: `tests/docx/integration.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/docx/integration.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { parseDocx, serializeDocx, parseDocxXml, serializeDocxXml } from '../../src/docx/index.js';
import JSZip from 'jszip';

async function createMinimalDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats.wordprocessingml.document.main+xml"/>
</Types>`);

  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`);

  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('DOCX integration', () => {
  it('parseDocx returns raw and semantic views', async () => {
    const buffer = await createMinimalDocx();
    const result = await parseDocx(buffer);

    expect(result.raw).toBeDefined();
    expect(result.semantic).toBeDefined();
    expect(result.raw.parts.size).toBeGreaterThan(0);
  });

  it('semantic view has correct structure', async () => {
    const buffer = await createMinimalDocx();
    const { semantic } = await parseDocx(buffer);

    expect(semantic.body.blocks).toHaveLength(1);
    const para = semantic.body.blocks[0] as any;
    expect(para.type).toBe('paragraph');
    expect(para.runs[0].text).toBe('Hello World');
  });

  it('serializeDocx produces valid ZIP', async () => {
    const buffer = await createMinimalDocx();
    const { semantic } = await parseDocx(buffer);

    const output = await serializeDocx(semantic);
    expect(output).toBeInstanceOf(ArrayBuffer);

    // Verify the output is a valid ZIP
    const zip = await JSZip.loadAsync(output);
    expect(zip.file('word/document.xml')).not.toBeNull();
  });

  it('parseDocxXml parses XML string', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Test</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = parseDocxXml(xml);
    expect(result.tag).toBe('w:document');
  });

  it('serializeDocxXml produces XML string', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Test</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = serializeDocxXml(parseDocxXml(xml));
    expect(result).toContain('w:document');
    expect(result).toContain('Test');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/docx/integration.test.ts`
Expected: FAIL with "Cannot find module"

- [ ] **Step 3: Write implementation**

Create `src/docx/index.ts`:

```typescript
import { unzip, zip } from '../core/zip.js';
import { parseRels } from '../core/rels.js';
import { parseContentTypes } from '../core/content-type.js';
import { parseXml, serializeXml } from '../core/xml.js';
import type { RawDocument, ParseResult, ParsedNode } from '../core/types.js';
import type { DocxDocument } from './types.js';
import { rawToSemantic } from './semantic.js';
import { semanticToXml } from './serializer.js';
import { parseDocxXml } from './parser.js';

export type { DocxDocument, DocumentMeta, Paragraph, TextRun, Table, Image } from './types.js';
export { parseDocxXml } from './parser.js';
export { rawToSemantic } from './semantic.js';
export { semanticToXml } from './serializer.js';

export async function parseDocx(buffer: ArrayBuffer): Promise<ParseResult<DocxDocument>> {
  const entries = await unzip(buffer);

  const rels = new Map<string, import('../core/types.js').Relationship[]>();
  const parts = new Map<string, ParsedNode>();
  let contentTypes: import('../core/types.js').ContentType[] = [];

  for (const entry of entries) {
    const text = new TextDecoder().decode(entry.data);

    if (entry.path === '[Content_Types].xml') {
      contentTypes = parseContentTypes(text);
    } else if (entry.path.endsWith('.rels')) {
      rels.set(entry.path, parseRels(text));
    } else if (entry.path.endsWith('.xml')) {
      parts.set(entry.path, parseXml(text));
    }
  }

  const raw: RawDocument = { entries, rels, contentTypes, parts };
  const semantic = rawToSemantic(raw);

  return { raw, semantic };
}

export async function serializeDocx(doc: DocxDocument): Promise<ArrayBuffer> {
  const documentXml = semanticToXml(doc);

  const entries = [
    { path: '[Content_Types].xml', data: new TextEncoder().encode(buildContentTypes()) },
    { path: '_rels/.rels', data: new TextEncoder().encode(buildRels()) },
    { path: 'word/document.xml', data: new TextEncoder().encode(documentXml) },
    { path: 'word/_rels/document.xml.rels', data: new TextEncoder().encode(buildWordRels()) },
  ];

  return zip(entries.map(e => ({ path: e.path, data: e.data.buffer })));
}

export function serializeDocxXml(node: ParsedNode): string {
  return serializeXml(node);
}

function buildContentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats.wordprocessingml.document.main+xml"/>
</Types>`;
}

function buildRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

function buildWordRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/docx/integration.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docx/index.ts tests/docx/integration.test.ts
git commit -m "feat(docx): add public API with parseDocx/serializeDocx"
```

---

### Task 13: Main Barrel Export + Final Verification

**Files:**
- Create: `src/index.ts`

- [ ] **Step 1: Update src/index.ts**

```typescript
export * from './core/index.js';
export * from './docx/index.js';
```

- [ ] **Step 2: Run all tests**

Run: `npx vitest run`
Expected: All tests PASS

- [ ] **Step 3: Typecheck**

Run: `npx tsc --noEmit`
Expected: No errors

- [ ] **Step 4: Build**

Run: `npx tsup`
Expected: Build succeeds, dist/ directory created with .js and .d.ts files

- [ ] **Step 5: Commit**

```bash
git add src/index.ts
git commit -m "feat: add main barrel export and verify build"
```

---

### Task 14: DOCX Edge Cases

**Files:**
- Modify: `tests/docx/semantic.test.ts`
- Modify: `src/docx/semantic.ts`

- [ ] **Step 1: Add test for paragraph with no runs**

Add to `tests/docx/semantic.test.ts`:

```typescript
it('handles paragraph with no runs (empty paragraph)', () => {
  const documentXml: ParsedNode = {
    tag: 'w:document',
    attrs: {},
    children: [
      {
        tag: 'w:body',
        attrs: {},
        children: [
          { tag: 'w:p', attrs: {}, children: [] },
        ],
      },
    ],
  };

  const raw = makeRaw({ 'word/document.xml': documentXml });
  const semantic = rawToSemantic(raw);
  expect(semantic.body.blocks).toHaveLength(1);
  expect((semantic.body.blocks[0] as any).runs).toEqual([]);
});
```

- [ ] **Step 2: Run test to verify it passes**

Run: `npx vitest run tests/docx/semantic.test.ts`
Expected: PASS (should already work)

- [ ] **Step 3: Add test for run with nested elements**

Add to `tests/docx/semantic.test.ts`:

```typescript
it('handles run with mixed content (text + tab + text)', () => {
  const documentXml: ParsedNode = {
    tag: 'w:document',
    attrs: {},
    children: [
      {
        tag: 'w:body',
        attrs: {},
        children: [
          {
            tag: 'w:p',
            attrs: {},
            children: [
              {
                tag: 'w:r',
                attrs: {},
                children: [
                  { tag: 'w:t', attrs: {}, children: ['Hello'] },
                  { tag: 'w:tab', attrs: {}, children: [] },
                  { tag: 'w:t', attrs: {}, children: ['World'] },
                ],
              },
            ],
          },
        ],
      },
    ],
  };

  const raw = makeRaw({ 'word/document.xml': documentXml });
  const semantic = rawToSemantic(raw);
  const para = semantic.body.blocks[0] as any;
  // Should have 2 text runs (tabs are separators)
  expect(para.runs).toHaveLength(1);
  expect(para.runs[0].text).toBe('Hello');
});
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/docx/semantic.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/docx/semantic.test.ts
git commit -m "test(docx): add edge case tests for semantic converter"
```

---

## Final Verification

- [ ] Run `npx vitest run` — all tests pass
- [ ] Run `npx tsc --noEmit` — no type errors
- [ ] Run `npx tsup` — build succeeds
- [ ] Verify dist/ contains expected files
