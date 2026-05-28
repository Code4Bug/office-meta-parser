import { parseXml } from '../core/xml.js';
import type { ParsedNode } from '../core/types.js';

export function parsePptxXml(xml: string): ParsedNode {
  return parseXml(xml);
}
