import { describe, it, expect } from 'vitest';
import type {
  PptxPresentation,
  Slide,
  SlideElement,
  TextShape,
  ImageShape,
  Position,
  ShapeStyle,
} from '../../src/pptx/types.js';

describe('PPTX types', () => {
  it('PptxPresentation structure', () => {
    const pres: PptxPresentation = {
      meta: { title: 'Test' },
      slides: [],
      masters: [],
      layouts: [],
    };
    expect(pres.meta.title).toBe('Test');
    expect(pres.slides).toEqual([]);
  });

  it('Slide structure', () => {
    const slide: Slide = {
      elements: [],
    };
    expect(slide.elements).toEqual([]);
  });

  it('TextShape structure', () => {
    const shape: TextShape = {
      type: 'text',
      content: 'Hello',
      position: { x: 100, y: 200, width: 300, height: 400 },
      paragraphs: [{ type: 'paragraph', runs: [{ text: 'Hello' }] }],
    };
    expect(shape.type).toBe('text');
    expect(shape.content).toBe('Hello');
    expect(shape.position.x).toBe(100);
  });

  it('ImageShape structure', () => {
    const shape: ImageShape = {
      type: 'image',
      relationshipId: 'rId1',
      position: { x: 0, y: 0, width: 100, height: 100 },
    };
    expect(shape.relationshipId).toBe('rId1');
  });

  it('Position structure', () => {
    const pos: Position = { x: 10, y: 20, width: 100, height: 50 };
    expect(pos.x).toBe(10);
    expect(pos.width).toBe(100);
  });

  it('ShapeStyle structure', () => {
    const style: ShapeStyle = {
      fill: { color: 'FF0000' },
      border: { color: '000000', width: 1 },
    };
    expect(style.fill?.color).toBe('FF0000');
  });
});
