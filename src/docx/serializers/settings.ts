import { serializeXml } from '../../core/xml.js';
import type { ParsedNode } from '../../core/types.js';
import type { DocumentSettings } from '../types.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function serializeSettings(settings: DocumentSettings): string {
  const children: ParsedNode[] = [];

  if (settings.zoom !== undefined) {
    children.push({ tag: 'w:zoom', attrs: { 'w:percent': String(settings.zoom) }, children: [] });
  }
  if (settings.defaultTabStop !== undefined) {
    children.push({ tag: 'w:defaultTabStop', attrs: { 'w:val': String(settings.defaultTabStop) }, children: [] });
  }
  if (settings.evenAndOddHeaders) {
    children.push({ tag: 'w:evenAndOddHeaders', attrs: { 'w:val': '1' }, children: [] });
  }
  if (settings.documentProtection) {
    children.push({ tag: 'w:documentProtection', attrs: { 'w:enforcement': '1' }, children: [] });
  }
  if (settings.characterSpacingControl) {
    children.push({ tag: 'w:characterSpacingControl', attrs: { 'w:val': settings.characterSpacingControl }, children: [] });
  }
  if (settings.compatibilityMode !== undefined) {
    children.push({
      tag: 'w:compat',
      attrs: {},
      children: [{
        tag: 'w:compatSetting',
        attrs: {
          'w:name': 'compatibilityMode',
          'w:uri': 'http://schemas.microsoft.com/office/word',
          'w:val': String(settings.compatibilityMode),
        },
        children: [],
      }],
    });
  }

  const root: ParsedNode = {
    tag: 'w:settings',
    attrs: { 'xmlns:w': W_NS },
    children,
  };

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializeXml(root);
}
