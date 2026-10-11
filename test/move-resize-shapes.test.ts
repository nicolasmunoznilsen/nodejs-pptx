import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/index.js';

describe('moving and resizing existing shapes', () => {
  it('gets and sets position and size in EMU', () => {
    const slide = Presentation.create().slides.add();
    slide.addTextBox('Hello', { x: 1, y: 2, width: 3, height: 1.5, name: 'Box' });
    const box = slide.shapes.get(0);
    expect(box.x).toBe(914400);
    expect(box.y).toBe(1828800);
    expect(box.width).toBe(2743200);
    expect(box.height).toBe(1371600);
    box.x = 0;
    box.y = 500000;
    box.width = 1000000;
    box.height = 250000;
    expect(box).toMatchObject({ x: 0, y: 500000, width: 1000000, height: 250000 });
  });

  it('moveTo and resize keep the other property untouched and chain', () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Hi', { x: 10, y: 10, width: 4, height: 2 });
    const box = slide.shapes.get(0);
    box.moveTo(100000, 200000);
    expect(box).toMatchObject({ x: 100000, y: 200000, width: 3657600, height: 1828800 });
    box.resize(500000, 300000);
    expect(box).toMatchObject({ x: 100000, y: 200000, width: 500000, height: 300000 });
    expect(box.moveTo(1, 2).resize(3, 4)).toBe(box);
  });

  it('creates a valid xfrm on shapes that lack one', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Hi');
    const shape = slide.shapes.get(0);
    delete shape.node['p:spPr']['a:xfrm'];
    shape.moveTo(123456, 654321);
    shape.resize(111111, 222222);
    expect(shape).toMatchObject({ x: 123456, y: 654321, width: 111111, height: 222222 });
    // a:xfrm must be the first child of p:spPr.
    expect(Object.keys(shape.node['p:spPr'])[0]).toBe('a:xfrm');
    const reopened = await Presentation.open(await pptx.toBuffer());
    const box = reopened.slides.get(0).shapes.get(0);
    expect(box).toMatchObject({ x: 123456, y: 654321, width: 111111, height: 222222 });
  });

  it('rejects invalid values without changing the shape', () => {
    const slide = Presentation.create().slides.add();
    slide.addTextBox('Hi', { x: 1, y: 1, width: 2, height: 2 });
    const box = slide.shapes.get(0);
    expect(() => { box.x = Number.NaN; }).toThrow(TypeError);
    expect(() => { box.y = Infinity; }).toThrow(TypeError);
    expect(() => { box.width = -1; }).toThrow(RangeError);
    expect(() => { box.height = 0; }).toThrow(RangeError);
    expect(() => box.resize(-5, 10)).toThrow(/width/);
    expect(box).toMatchObject({ x: 914400, y: 914400, width: 1828800, height: 1828800 });
  });

  it('persists moved and resized shapes through save and reopen', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Hi', { x: 1, y: 1, width: 2, height: 1, name: 'Box' });
    const box = slide.shapes.findByName('Box')!;
    box.moveTo(200000, 300000).resize(400000, 500000);
    const reopened = await Presentation.open(await pptx.toBuffer());
    const saved = reopened.slides.get(0).shapes.get(0);
    expect(saved).toMatchObject({ x: 200000, y: 300000, width: 400000, height: 500000 });
    expect(reopened.slides.get(0).text).toBe('Hi');
  });
});
