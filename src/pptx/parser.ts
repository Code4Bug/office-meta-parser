import { parseXml } from '../core/xml.js';
import type { ParsedNode } from '../core/types.js';

export function parsePptxXml(xml: string): ParsedNode {
  const root = parseXml(xml);
  stripWhitespace(root);
  return root;
}

function stripWhitespace(node: ParsedNode): void {
  node.children = node.children.filter(
    (child) => typeof child !== 'string' || child.trim() !== '',
  );
  for (const child of node.children) {
    if (typeof child !== 'string') {
      stripWhitespace(child);
    }
  }
}
