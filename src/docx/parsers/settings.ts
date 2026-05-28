import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { DocumentSettings } from '../types.js';
import { findChild } from './utils.js';

export function extractSettings(raw: RawDocument): DocumentSettings | undefined {
  const settingsXml = raw.parts.get('word/settings.xml');
  if (!settingsXml) return undefined;

  const settings: DocumentSettings = {};

  for (const child of settingsXml.children) {
    if (typeof child === 'string') continue;
    switch (child.tag) {
      case 'w:defaultTabStop':
        settings.defaultTabStop = parseInt(child.attrs['w:val'] || '0', 10);
        break;
      case 'w:zoom':
        settings.zoom = parseInt(child.attrs['w:percent'] || '100', 10);
        break;
      case 'w:evenAndOddHeaders':
        settings.evenAndOddHeaders = child.attrs['w:val'] === '1' || child.attrs['w:val'] === 'true';
        break;
      case 'w:documentProtection':
        settings.documentProtection = child.attrs['w:enforcement'] === '1' || child.attrs['w:enforcement'] === 'true';
        break;
      case 'w:characterSpacingControl':
        settings.characterSpacingControl = child.attrs['w:val'];
        break;
      case 'w:compat': {
        const compatMode = findCompatMode(child);
        if (compatMode !== undefined) settings.compatibilityMode = compatMode;
        break;
      }
    }
  }

  return Object.keys(settings).length > 0 ? settings : undefined;
}

function findCompatMode(node: ParsedNode): number | undefined {
  for (const child of node.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:compatSetting') {
      const name = child.attrs['w:name'];
      if (name === 'compatibilityMode') {
        const val = child.attrs['w:val'];
        if (val) return parseInt(val, 10);
      }
    }
  }
  return undefined;
}
