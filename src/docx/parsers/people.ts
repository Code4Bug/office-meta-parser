import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { Person } from '../types.js';
import { findChild } from './utils.js';

export function extractPeople(raw: RawDocument): Person[] | undefined {
  const xml = raw.parts.get('word/people.xml');
  if (!xml) return undefined;

  const result: Person[] = [];

  for (const child of xml.children) {
    if (typeof child === 'string') continue;
    if (child.tag !== 'w15:person') continue;

    const author = child.attrs['w15:author'];
    if (!author) continue;

    const person: Person = { author };

    const presenceInfo = findChild(child, 'w15:presenceInfo');
    if (presenceInfo) {
      if (presenceInfo.attrs['w15:userId']) person.userId = presenceInfo.attrs['w15:userId'];
      if (presenceInfo.attrs['w15:providerId']) person.providerId = presenceInfo.attrs['w15:providerId'];
    }

    result.push(person);
  }

  return result.length > 0 ? result : undefined;
}
