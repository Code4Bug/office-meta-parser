import { serializeXml } from './xml.js';
import type { ParsedNode } from './types.js';

export type CustomPropertyType = 'lpwstr' | 'i4' | 'r8' | 'date' | 'bool' | 'blob';

export interface CustomProperty {
  name: string;
  value: string | number | boolean;
  type: CustomPropertyType;
}

const CP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

function getDirectText(node: ParsedNode): string {
  for (const child of node.children) {
    if (typeof child === 'string') return child;
  }
  return '';
}

export function parseCustomProperties(node: ParsedNode): CustomProperty[] {
  const props: CustomProperty[] = [];

  for (const child of node.children) {
    if (typeof child === 'string' || child.tag !== 'property') continue;

    const name = child.attrs['name'];
    if (!name) continue;

    const valueNode = child.children.find(c => typeof c !== 'string') as ParsedNode | undefined;
    if (!valueNode) continue;

    const text = getDirectText(valueNode);
    const typeTag = valueNode.tag.replace('vt:', '') as CustomPropertyType;

    let value: string | number | boolean = text;
    if (typeTag === 'i4' || typeTag === 'r8') {
      value = parseFloat(text);
    } else if (typeTag === 'bool') {
      value = text === 'true' || text === '1';
    }

    props.push({ name, value, type: typeTag });
  }

  return props;
}

export function serializeCustomProperties(props: CustomProperty[]): string {
  const children: ParsedNode[] = props.map((prop, i) => {
    const valueChildren: ParsedNode[] = [];

    const vtTag = `vt:${prop.type}`;
    const textValue = prop.type === 'bool'
      ? (prop.value ? 'true' : 'false')
      : String(prop.value);

    valueChildren.push({ tag: vtTag, attrs: {}, children: [textValue] });

    return {
      tag: 'property',
      attrs: {
        fmtid: '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
        pid: String(i + 2),
        name: prop.name,
      },
      children: valueChildren,
    };
  });

  const root: ParsedNode = {
    tag: 'Properties',
    attrs: {
      xmlns: CP_NS,
      'xmlns:vt': VT_NS,
    },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
