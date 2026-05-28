import { describe, it, expect } from 'vitest';
import { extractRevisions } from '../../src/docx/parsers/revision.js';
import type { RawDocument, ParsedNode } from '../../src/core/types.js';

function makeRaw(parts: Record<string, ParsedNode>): RawDocument {
  return {
    entries: [],
    rels: new Map(),
    contentTypes: [],
    parts: new Map(Object.entries(parts)),
  };
}

describe('extractRevisions', () => {
  it('returns undefined when no document.xml', () => {
    const raw = makeRaw({});
    expect(extractRevisions(raw)).toBeUndefined();
  });

  it('returns undefined when no revisions', () => {
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
                { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Hello'] }] },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    expect(extractRevisions(raw)).toBeUndefined();
  });

  it('parses insert revision', () => {
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
                  tag: 'w:ins',
                  attrs: { 'w:id': '1', 'w:author': 'John', 'w:date': '2023-01-01T00:00:00Z' },
                  children: [
                    { tag: 'w:r', attrs: {}, children: [{ tag: 'w:t', attrs: {}, children: ['Added text'] }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const revisions = extractRevisions(raw);
    expect(revisions).toHaveLength(1);
    expect(revisions![0].type).toBe('insert');
    expect(revisions![0].author).toBe('John');
    expect(revisions![0].date).toBe('2023-01-01T00:00:00Z');
  });

  it('parses delete revision', () => {
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
                  tag: 'w:del',
                  attrs: { 'w:id': '2', 'w:author': 'Jane', 'w:date': '2023-02-01T00:00:00Z' },
                  children: [
                    { tag: 'w:r', attrs: {}, children: [{ tag: 'w:delText', attrs: {}, children: ['Deleted text'] }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const revisions = extractRevisions(raw);
    expect(revisions).toHaveLength(1);
    expect(revisions![0].type).toBe('delete');
    expect(revisions![0].author).toBe('Jane');
  });

  it('parses format change revision', () => {
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
                        {
                          tag: 'w:rPrChange',
                          attrs: { 'w:id': '3', 'w:author': 'Bob', 'w:date': '2023-03-01T00:00:00Z' },
                          children: [],
                        },
                      ],
                    },
                    { tag: 'w:t', attrs: {}, children: ['Formatted'] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const revisions = extractRevisions(raw);
    expect(revisions).toHaveLength(1);
    expect(revisions![0].type).toBe('formatChange');
    expect(revisions![0].author).toBe('Bob');
  });

  it('parses multiple revisions', () => {
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
                  tag: 'w:ins',
                  attrs: { 'w:id': '1', 'w:author': 'John', 'w:date': '2023-01-01T00:00:00Z' },
                  children: [],
                },
                {
                  tag: 'w:del',
                  attrs: { 'w:id': '2', 'w:author': 'Jane', 'w:date': '2023-02-01T00:00:00Z' },
                  children: [],
                },
              ],
            },
          ],
        },
      ],
    };

    const raw = makeRaw({ 'word/document.xml': documentXml });
    const revisions = extractRevisions(raw);
    expect(revisions).toHaveLength(2);
    expect(revisions![0].type).toBe('insert');
    expect(revisions![1].type).toBe('delete');
  });
});
