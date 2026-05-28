import { describe, it, expect } from 'vitest';
import { serializeCellStyles } from '../../../src/xlsx/serializers/styles.js';
import type { CellStyleDefinitions } from '../../../src/xlsx/types.js';

function makeStyles(overrides: Partial<CellStyleDefinitions> = {}): CellStyleDefinitions {
  return {
    cellStyles: [],
    fonts: [],
    fills: [],
    borders: [],
    numberFormats: [],
    ...overrides,
  };
}

describe('XLSX P0 serializer - styles', () => {
  it('serializes font vertAlign', () => {
    const xml = serializeCellStyles(makeStyles({
      fonts: [{ vertAlign: 'superscript', size: 12 }],
    }));

    expect(xml).toContain('<vertAlign val="superscript"/>');
  });

  it('serializes font scheme', () => {
    const xml = serializeCellStyles(makeStyles({
      fonts: [{ scheme: 'major' }],
    }));

    expect(xml).toContain('<scheme val="major"/>');
  });

  it('serializes gradient fill', () => {
    const xml = serializeCellStyles(makeStyles({
      fills: [{
        gradientFill: {
          type: 'linear',
          degree: 90,
          stops: [
            { position: 0, color: 'FF0000' },
            { position: 1, color: '0000FF' },
          ],
        },
      }],
    }));

    expect(xml).toContain('<gradientFill');
    expect(xml).toContain('type="linear"');
    expect(xml).toContain('degree="90"');
    expect(xml).toContain('<stop position="0">');
    expect(xml).toContain('rgb="FF0000"');
    expect(xml).toContain('rgb="0000FF"');
  });

  it('serializes diagonal border', () => {
    const xml = serializeCellStyles(makeStyles({
      borders: [{
        diagonal: { style: 'thin', color: 'FF0000' },
        diagonalUp: true,
        diagonalDown: true,
      }],
    }));

    expect(xml).toContain('<diagonal style="thin">');
    expect(xml).toContain('diagonalUp="1"');
    expect(xml).toContain('diagonalDown="1"');
  });

  it('serializes alignment indent, textRotation, shrinkToFit', () => {
    const xml = serializeCellStyles(makeStyles({
      cellStyles: [{
        id: '0',
        alignment: { horizontal: 'left', indent: 2, textRotation: 45, shrinkToFit: true },
      }],
    }));

    expect(xml).toContain('indent="2"');
    expect(xml).toContain('textRotation="45"');
    expect(xml).toContain('shrinkToFit="1"');
  });

  it('serializes combined font properties', () => {
    const xml = serializeCellStyles(makeStyles({
      fonts: [{
        name: 'Calibri',
        size: 11,
        bold: true,
        vertAlign: 'subscript',
        color: 'FF0000',
        scheme: 'minor',
      }],
    }));

    expect(xml).toContain('<name val="Calibri"/>');
    expect(xml).toContain('<vertAlign val="subscript"/>');
    expect(xml).toContain('<scheme val="minor"/>');
    expect(xml).toContain('<color rgb="FF0000"/>');
  });
});
