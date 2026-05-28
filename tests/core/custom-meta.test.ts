import { describe, it, expect } from 'vitest';
import { parseCustomProperties, serializeCustomProperties } from '../../src/core/custom-meta.js';
import { parseXml } from '../../src/core/xml.js';
import type { CustomProperty } from '../../src/core/custom-meta.js';

describe('custom-meta', () => {
  describe('parseCustomProperties', () => {
    it('parses empty Properties', () => {
      const xml = '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>';
      const node = parseXml(xml);
      const props = parseCustomProperties(node);
      expect(props).toEqual([]);
    });

    it('parses string property', () => {
      const xml = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Editor">
    <vt:lpwstr>John Doe</vt:lpwstr>
  </property>
</Properties>`;
      const node = parseXml(xml);
      const props = parseCustomProperties(node);
      expect(props).toHaveLength(1);
      expect(props[0].name).toBe('Editor');
      expect(props[0].value).toBe('John Doe');
      expect(props[0].type).toBe('lpwstr');
    });

    it('parses number property', () => {
      const xml = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Count">
    <vt:i4>42</vt:i4>
  </property>
</Properties>`;
      const node = parseXml(xml);
      const props = parseCustomProperties(node);
      expect(props).toHaveLength(1);
      expect(props[0].name).toBe('Count');
      expect(props[0].value).toBe(42);
      expect(props[0].type).toBe('i4');
    });

    it('parses boolean property', () => {
      const xml = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Draft">
    <vt:bool>true</vt:bool>
  </property>
</Properties>`;
      const node = parseXml(xml);
      const props = parseCustomProperties(node);
      expect(props).toHaveLength(1);
      expect(props[0].name).toBe('Draft');
      expect(props[0].value).toBe(true);
      expect(props[0].type).toBe('bool');
    });

    it('parses multiple properties', () => {
      const xml = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Status">
    <vt:lpwstr>Active</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Priority">
    <vt:i4>1</vt:i4>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Reviewed">
    <vt:bool>false</vt:bool>
  </property>
</Properties>`;
      const node = parseXml(xml);
      const props = parseCustomProperties(node);
      expect(props).toHaveLength(3);
      expect(props[0].name).toBe('Status');
      expect(props[0].value).toBe('Active');
      expect(props[1].name).toBe('Priority');
      expect(props[1].value).toBe(1);
      expect(props[2].name).toBe('Reviewed');
      expect(props[2].value).toBe(false);
    });

    it('parses float property', () => {
      const xml = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Score">
    <vt:r8>98.5</vt:r8>
  </property>
</Properties>`;
      const node = parseXml(xml);
      const props = parseCustomProperties(node);
      expect(props).toHaveLength(1);
      expect(props[0].name).toBe('Score');
      expect(props[0].value).toBe(98.5);
      expect(props[0].type).toBe('r8');
    });
  });

  describe('serializeCustomProperties', () => {
    it('serializes empty array', () => {
      const xml = serializeCustomProperties([]);
      expect(xml).toContain('Properties');
      expect(xml).toContain('xmlns=');
    });

    it('serializes string property', () => {
      const xml = serializeCustomProperties([{ name: 'Author', value: 'Jane', type: 'lpwstr' }]);
      expect(xml).toContain('name="Author"');
      expect(xml).toContain('<vt:lpwstr>Jane</vt:lpwstr>');
    });

    it('serializes number property', () => {
      const xml = serializeCustomProperties([{ name: 'Count', value: 42, type: 'i4' }]);
      expect(xml).toContain('name="Count"');
      expect(xml).toContain('<vt:i4>42</vt:i4>');
    });

    it('serializes boolean property', () => {
      const xml = serializeCustomProperties([{ name: 'Draft', value: true, type: 'bool' }]);
      expect(xml).toContain('name="Draft"');
      expect(xml).toContain('<vt:bool>true</vt:bool>');
    });

    it('serializes boolean false', () => {
      const xml = serializeCustomProperties([{ name: 'Final', value: false, type: 'bool' }]);
      expect(xml).toContain('<vt:bool>false</vt:bool>');
    });

    it('serializes multiple properties with sequential pids', () => {
      const props: CustomProperty[] = [
        { name: 'A', value: 'x', type: 'lpwstr' },
        { name: 'B', value: 1, type: 'i4' },
      ];
      const xml = serializeCustomProperties(props);
      expect(xml).toContain('pid="2"');
      expect(xml).toContain('pid="3"');
    });

    it('serializes float property', () => {
      const xml = serializeCustomProperties([{ name: 'Rate', value: 3.14, type: 'r8' }]);
      expect(xml).toContain('<vt:r8>3.14</vt:r8>');
    });
  });

  describe('round-trip', () => {
    it('parse then serialize preserves data', () => {
      const original: CustomProperty[] = [
        { name: 'Project', value: 'Parser', type: 'lpwstr' },
        { name: 'Version', value: 2, type: 'i4' },
        { name: 'Active', value: true, type: 'bool' },
        { name: 'Score', value: 99.9, type: 'r8' },
      ];

      const xml = serializeCustomProperties(original);
      const node = parseXml(xml);
      const parsed = parseCustomProperties(node);

      expect(parsed).toHaveLength(4);
      expect(parsed[0].name).toBe('Project');
      expect(parsed[0].value).toBe('Parser');
      expect(parsed[0].type).toBe('lpwstr');
      expect(parsed[1].name).toBe('Version');
      expect(parsed[1].value).toBe(2);
      expect(parsed[1].type).toBe('i4');
      expect(parsed[2].name).toBe('Active');
      expect(parsed[2].value).toBe(true);
      expect(parsed[2].type).toBe('bool');
      expect(parsed[3].name).toBe('Score');
      expect(parsed[3].value).toBeCloseTo(99.9);
      expect(parsed[3].type).toBe('r8');
    });
  });
});
