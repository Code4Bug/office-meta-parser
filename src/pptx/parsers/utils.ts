import type { ParsedNode } from '../../core/types.js';

export function findChild(node: ParsedNode, tag: string): ParsedNode | undefined {
  for (const child of node.children) {
    if (typeof child !== 'string' && child.tag === tag) {
      return child;
    }
  }
  return undefined;
}

/** Match both 'tag' and 'ns:tag' for namespace-prefixed elements */
export function findChildAny(node: ParsedNode, tag: string): ParsedNode | undefined {
  for (const child of node.children) {
    if (typeof child !== 'string' && (child.tag === tag || child.tag.endsWith(':' + tag))) {
      return child;
    }
  }
  return undefined;
}

export function getTextContent(node: ParsedNode): string {
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
