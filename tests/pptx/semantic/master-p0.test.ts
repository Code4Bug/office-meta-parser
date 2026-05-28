import { describe, it, expect } from 'vitest';
import type { ParsedNode, RawDocument, Relationship } from '../../../src/core/types.js';
import { rawToSemantic } from '../../../src/pptx/semantic.js';

function makeNode(tag: string, attrs: Record<string, string> = {}, children: (ParsedNode | string)[] = []): ParsedNode {
  return { tag, attrs, children };
}

function makeRaw(parts: Record<string, ParsedNode>, rels?: Map<string, Relationship[]>): RawDocument {
  return {
    entries: [],
    rels: rels || new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('PPTX P0 semantic - master features', () => {
  it('parses txStyles from slide master', () => {
    const masterXml = makeNode('p:sldMaster', {}, [
      makeNode('p:cSld', {}, [
        makeNode('p:bg', {}, [
          makeNode('p:bgPr', {}, [
            makeNode('a:solidFill', {}, [makeNode('a:srgbClr', { val: 'FFFFFF' })]),
          ]),
        ]),
        makeNode('p:spTree'),
      ]),
      makeNode('p:sldLayoutIdLst'),
      makeNode('p:txStyles', {}, [
        makeNode('p:titleStyle', {}, [
          makeNode('a:lvl1pPr', { algn: 'ctr', marL: '0', indent: '0' }, [
            makeNode('a:spcBef', {}, [makeNode('a:spcPts', { val: '0' })]),
            makeNode('a:spcAft', {}, [makeNode('a:spcPts', { val: '800' })]),
            makeNode('a:defRPr', { sz: '3600' }),
          ]),
        ]),
        makeNode('p:bodyStyle', {}, [
          makeNode('a:lvl1pPr', { algn: 'l' }, [
            makeNode('a:defRPr', { sz: '1800' }),
          ]),
        ]),
        makeNode('p:otherStyle', {}, [
          makeNode('a:lvl1pPr', { algn: 'l' }),
        ]),
      ]),
    ]);

    const raw = makeRaw({ 'ppt/slideMasters/slideMaster1.xml': masterXml });
    const sem = rawToSemantic(raw);
    expect(sem.masters).toHaveLength(1);
    expect(sem.masters[0].txStyles).toBeDefined();
    expect(sem.masters[0].txStyles!.titleStyle).toBeDefined();
    expect(sem.masters[0].txStyles!.titleStyle!.levels).toHaveLength(1);
    expect(sem.masters[0].txStyles!.titleStyle!.levels[0].level).toBe(1);
    expect(sem.masters[0].txStyles!.titleStyle!.levels[0].alignment).toBe('ctr');
    expect(sem.masters[0].txStyles!.titleStyle!.levels[0].fontScale).toBe(36);
    expect(sem.masters[0].txStyles!.titleStyle!.levels[0].spcAft).toBe(8);
    expect(sem.masters[0].txStyles!.bodyStyle).toBeDefined();
    expect(sem.masters[0].txStyles!.bodyStyle!.levels[0].fontScale).toBe(18);
    expect(sem.masters[0].txStyles!.otherStyle).toBeDefined();
  });
});
