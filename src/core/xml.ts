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
    (item: any) => {
      if (typeof item !== 'object' || '__text' in item || '__comment' in item) return false;
      // Skip processing instructions like <?xml ...?>
      const tag = Object.keys(item).find(k => k !== ':@');
      return tag != null && !tag.startsWith('?');
    }
  );
  if (!root) {
    throw new Error('No root element found in XML');
  }

  return convertFromPreserveOrder(root);
}

function convertFromPreserveOrder(raw: any): ParsedNode {
  // In preserveOrder mode, element is { tagName: [...children], ':@': { ...attrs } }
  const tag = Object.keys(raw).find(k => k !== ':@') || '';
  const children: (ParsedNode | string)[] = [];

  const rawChildren = raw[tag];
  if (Array.isArray(rawChildren)) {
    for (const item of rawChildren) {
      if ('__text' in item) {
        children.push(String(item.__text));
      } else if ('__comment' in item) {
        // Skip comments
      } else {
        children.push(convertFromPreserveOrder(item));
      }
    }
  }

  const attrs: Record<string, string> = {};
  if (raw[':@']) {
    for (const [key, value] of Object.entries(raw[':@'])) {
      attrs[key] = String(value);
    }
  }

  return { tag, attrs, children };
}

export function serializeXml(node: ParsedNode): string {
  const raw = convertToPreserveOrder(node);
  const builder = new XMLBuilder(builderOptions);
  return builder.build([raw]);
}

function convertToPreserveOrder(node: ParsedNode): any {
  const result: any = {};
  result[node.tag] = [];

  // Add children
  for (const child of node.children) {
    if (typeof child === 'string') {
      result[node.tag].push({ __text: child });
    } else {
      result[node.tag].push(convertToPreserveOrder(child));
    }
  }

  // Add attributes
  if (Object.keys(node.attrs).length > 0) {
    result[':@'] = { ...node.attrs };
  }

  return result;
}
