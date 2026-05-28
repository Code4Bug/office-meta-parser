import { describe, it, expect } from 'vitest';
import { serializePeople } from '../../../src/docx/serializer.js';
import type { Person } from '../../../src/docx/types.js';

describe('DOCX serializer - people', () => {
  it('serializes people with presence info', () => {
    const people: Person[] = [
      { author: '张三', userId: '12345', providerId: 'WPS Office' },
      { author: '李四' },
    ];

    const xml = serializePeople(people);
    expect(xml).toContain('w15:people');
    expect(xml).toContain('w15:person');
    expect(xml).toContain('w15:author="张三"');
    expect(xml).toContain('w15:author="李四"');
    expect(xml).toContain('w15:presenceInfo');
    expect(xml).toContain('w15:userId="12345"');
    expect(xml).toContain('w15:providerId="WPS Office"');
  });

  it('omits presenceInfo when no userId/providerId', () => {
    const people: Person[] = [{ author: '王五' }];
    const xml = serializePeople(people);
    expect(xml).toContain('w15:author="王五"');
    expect(xml).not.toContain('w15:presenceInfo');
  });
});
