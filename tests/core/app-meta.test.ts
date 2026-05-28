import { describe, it, expect } from 'vitest';
import { parseAppMeta, serializeAppMeta } from '../../src/core/app-meta.js';
import { parseXml } from '../../src/core/xml.js';
import type { AppMeta } from '../../src/core/app-meta.js';

describe('app-meta', () => {
  describe('parseAppMeta', () => {
    it('parses empty Properties', () => {
      const xml = '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"></Properties>';
      const node = parseXml(xml);
      const meta = parseAppMeta(node);
      expect(meta).toEqual({});
    });

    it('parses all fields', () => {
      const xml = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Template>Normal.dotm</Template>
  <TotalTime>15</TotalTime>
  <Pages>3</Pages>
  <Words>500</Words>
  <Characters>2500</Characters>
  <CharactersWithSpaces>3000</CharactersWithSpaces>
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>Acme Corp</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0.12345</AppVersion>
</Properties>`;
      const node = parseXml(xml);
      const meta = parseAppMeta(node);
      expect(meta.template).toBe('Normal.dotm');
      expect(meta.totalTime).toBe(15);
      expect(meta.pages).toBe(3);
      expect(meta.words).toBe(500);
      expect(meta.characters).toBe(2500);
      expect(meta.charactersWithSpaces).toBe(3000);
      expect(meta.application).toBe('Microsoft Office Word');
      expect(meta.docSecurity).toBe(0);
      expect(meta.scaleCrop).toBe(false);
      expect(meta.company).toBe('Acme Corp');
      expect(meta.linksUpToDate).toBe(false);
      expect(meta.sharedDoc).toBe(false);
      expect(meta.hyperlinksChanged).toBe(false);
      expect(meta.appVersion).toBe('16.0.12345');
    });

    it('parses boolean true values', () => {
      const xml = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <ScaleCrop>true</ScaleCrop>
  <LinksUpToDate>true</LinksUpToDate>
  <SharedDoc>true</SharedDoc>
  <HyperlinksChanged>true</HyperlinksChanged>
</Properties>`;
      const node = parseXml(xml);
      const meta = parseAppMeta(node);
      expect(meta.scaleCrop).toBe(true);
      expect(meta.linksUpToDate).toBe(true);
      expect(meta.sharedDoc).toBe(true);
      expect(meta.hyperlinksChanged).toBe(true);
    });
  });

  describe('serializeAppMeta', () => {
    it('serializes empty meta', () => {
      const xml = serializeAppMeta({});
      expect(xml).toContain('Properties');
      expect(xml).toContain('xmlns=');
    });

    it('serializes template', () => {
      const xml = serializeAppMeta({ template: 'Normal.dotm' });
      expect(xml).toContain('<Template>Normal.dotm</Template>');
    });

    it('serializes numeric fields', () => {
      const xml = serializeAppMeta({ pages: 5, words: 1000, characters: 5000 });
      expect(xml).toContain('<Pages>5</Pages>');
      expect(xml).toContain('<Words>1000</Words>');
      expect(xml).toContain('<Characters>5000</Characters>');
    });

    it('serializes boolean fields', () => {
      const xml = serializeAppMeta({ scaleCrop: true, sharedDoc: false });
      expect(xml).toContain('<ScaleCrop>true</ScaleCrop>');
      expect(xml).toContain('<SharedDoc>false</SharedDoc>');
    });

    it('serializes application and version', () => {
      const xml = serializeAppMeta({ application: 'LibreOffice', appVersion: '7.5.2' });
      expect(xml).toContain('<Application>LibreOffice</Application>');
      expect(xml).toContain('<AppVersion>7.5.2</AppVersion>');
    });

    it('serializes complete meta', () => {
      const meta: AppMeta = {
        template: 'Custom.dotx',
        totalTime: 30,
        pages: 10,
        words: 2000,
        application: 'Microsoft Office Word',
        appVersion: '16.0000',
        company: 'Test Inc',
      };
      const xml = serializeAppMeta(meta);
      expect(xml).toContain('Custom.dotx');
      expect(xml).toContain('<TotalTime>30</TotalTime>');
      expect(xml).toContain('<Pages>10</Pages>');
      expect(xml).toContain('<Words>2000</Words>');
      expect(xml).toContain('Microsoft Office Word');
      expect(xml).toContain('Test Inc');
    });
  });

  describe('round-trip', () => {
    it('parse then serialize preserves data', () => {
      const original: AppMeta = {
        template: 'Report.dotx',
        totalTime: 42,
        pages: 7,
        words: 1500,
        characters: 8000,
        application: 'Office',
        appVersion: '16.0.12345',
        company: 'ACME',
        scaleCrop: true,
      };

      const xml = serializeAppMeta(original);
      const node = parseXml(xml);
      const parsed = parseAppMeta(node);

      expect(parsed.template).toBe(original.template);
      expect(parsed.totalTime).toBe(original.totalTime);
      expect(parsed.pages).toBe(original.pages);
      expect(parsed.words).toBe(original.words);
      expect(parsed.characters).toBe(original.characters);
      expect(parsed.application).toBe(original.application);
      expect(parsed.appVersion).toBe(original.appVersion);
      expect(parsed.company).toBe(original.company);
      expect(parsed.scaleCrop).toBe(original.scaleCrop);
    });
  });
});
