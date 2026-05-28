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
