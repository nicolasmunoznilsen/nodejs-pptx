import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Presentation, MARGIN_ATTR, MARGIN_DEFAULT_EMU, marginAttrFor, readMarginInches, writeMarginInches, assertMarginInches, assertMarginSide, type TextFrameMarginSide } from '../src/index.js';

async function slideXml(pptx: Presentation): Promise<string> {
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  return (await zip.file('ppt/slides/slide1.xml')!.async('string'));
}

const BODY_PR = '<a:bodyPr wrap="square"/>';

async function withSlideXml(mutate: (xml: string) => string): Promise<Presentation> {
  const pptx = Presentation.create();
  pptx.slides.add().addTextBox('external', { name: 'External' });
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  const path = 'ppt/slides/slide1.xml';
  const xml = (await zip.file(path)!.async('string')).replace('<a:bodyPr wrap="square"/>', BODY_PR);
  zip.file(path, mutate(xml));
  return Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
}

describe('text-frame margins', () => {
  const sides: Array<[TextFrameMarginSide, string, number]> = [
    ['left', 'lIns', 0.1],
    ['top', 'tIns', 0.05],
    ['right', 'rIns', 0.1],
    ['bottom', 'bIns', 0.05],
  ];

  it('reports the OOXML defaults when no inset is explicit', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('plain', { name: 'Plain' });
    const frame = pptx.slides.get(0).shapes.findByName('Plain')!.textFrame;
    for (const [side, , inches] of sides) {
      expect(frame[side === 'left' ? 'marginLeft' : side === 'top' ? 'marginTop' : side === 'right' ? 'marginRight' : 'marginBottom']).toBe(inches);
      expect(readMarginInches(undefined, side)).toBe(inches);
    }
    expect(frame.marginLeft).toBe(0.1);
    expect(frame.marginTop).toBe(0.05);
    expect(frame.marginRight).toBe(0.1);
    expect(frame.marginBottom).toBe(0.05);
    expect(frame.margin).toBeUndefined(); // 0.1 != 0.05
    expect(await slideXml(pptx)).not.toMatch(/Ins="/);
  });

  it('writes and reads every side through the text frame', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('margins', { name: 'Margins' });
    const frame = slide.shapes.findByName('Margins')!.textFrame;

    frame.marginLeft = 0.25;
    frame.marginTop = 0.2;
    frame.marginRight = 0.3;
    frame.marginBottom = 0.15;
    expect(frame.marginLeft).toBe(0.25);
    expect(frame.marginTop).toBe(0.2);
    expect(frame.marginRight).toBe(0.3);
    expect(frame.marginBottom).toBe(0.15);
    expect(frame.margin).toBeUndefined();

    const xml = await slideXml(pptx);
    expect(xml).toContain('lIns="228600"');
    expect(xml).toContain('tIns="182880"');
    expect(xml).toContain('rIns="274320"');
    expect(xml).toContain('bIns="137160"');

    const reopened = await Presentation.open(await pptx.toBuffer());
    const reopenedFrame = reopened.slides.get(0).shapes.findByName('Margins')!.textFrame;
    expect(reopenedFrame.marginLeft).toBe(0.25);
    expect(reopenedFrame.marginTop).toBe(0.2);
    expect(reopenedFrame.marginRight).toBe(0.3);
    expect(reopenedFrame.marginBottom).toBe(0.15);
  });

  it('clears a side back to the inherited default with undefined', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('clear', { name: 'Clear', marginLeft: 0.5 });
    const frame = slide.shapes.findByName('Clear')!.textFrame;
    expect(frame.marginLeft).toBe(0.5);
    frame.marginLeft = undefined;
    expect(frame.marginLeft).toBe(0.1);
    expect(await slideXml(pptx)).not.toContain('lIns=');
  });

  it('supports the margin convenience accessor', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('uniform', { name: 'Uniform' });
    const frame = slide.shapes.findByName('Uniform')!.textFrame;

    frame.margin = 0.4;
    expect(frame.margin).toBe(0.4);
    expect([frame.marginLeft, frame.marginTop, frame.marginRight, frame.marginBottom]).toEqual([0.4, 0.4, 0.4, 0.4]);
    const xml = await slideXml(pptx);
    for (const attr of ['lIns', 'tIns', 'rIns', 'bIns']) expect(xml).toContain(`${attr}="365760"`);

    frame.marginRight = 0.05;
    expect(frame.margin).toBeUndefined();

    frame.margin = undefined;
    expect(await slideXml(pptx)).not.toMatch(/Ins="/);
    expect(frame.margin).toBeUndefined();
  });

  it('keeps unrelated bodyPr attributes and children intact', async () => {
    const opened = await withSlideXml(xml => xml.replace(
      BODY_PR,
      '<a:bodyPr wrap="none" lIns="45720" anchor="ctr" rtlCol="0"><a:normAutofit fontScale="90000" lnSpcReduction="10000"/><a:sp3d/></a:bodyPr>',
    ));
    const frame = opened.slides.get(0).shapes.findByName('External')!.textFrame;
    expect(frame.marginLeft).toBe(0.05);
    expect(frame.marginTop).toBe(0.05); // inherited default

    frame.marginTop = 0.3;
    const xml = await slideXml(opened);
    expect(xml).toContain('wrap="none"');
    expect(xml).toContain('anchor="ctr"');
    expect(xml).toContain('rtlCol="0"');
    expect(xml).toContain('lIns="45720"');
    expect(xml).toContain('tIns="274320"');
    expect(xml).toContain('<a:normAutofit fontScale="90000" lnSpcReduction="10000"/>');
    expect(xml).toContain('<a:sp3d/>');
    expect(await opened.slides.get(0).text).toBe('external');
  });

  it('creates a:bodyPr as the first child of p:txBody, only on write', async () => {
    const opened = await withSlideXml(xml => xml.replace(BODY_PR, ''));
    const frame = opened.slides.get(0).shapes.findByName('External')!.textFrame;
    // Reading defaults must not create the node.
    expect(frame.marginLeft).toBe(0.1);
    expect(await slideXml(opened)).not.toContain('a:bodyPr');

    frame.marginBottom = 0.2;
    const xml = await slideXml(opened);
    expect(xml).toMatch(/<p:txBody><a:bodyPr bIns="182880"\/><a:lstStyle\/>/);
    const reopened = await Presentation.open(await opened.toBuffer());
    expect(reopened.slides.get(0).shapes.findByName('External')!.textFrame.marginBottom).toBe(0.2);
    expect(reopened.slides.get(0).text).toBe('external');
  });

  it('accepts margin options in addTextBox and reports them in getTextBoxes', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const uniform = slide.addTextBox('uniform', { name: 'Uniform', margin: 0.2 });
    const mixed = slide.addTextBox('mixed', { name: 'Mixed', marginLeft: 0.01, marginTop: 0.02 });
    const plain = slide.addTextBox('plain', { name: 'Plain' });

    expect(uniform).toMatchObject({ margin: 0.2, marginLeft: 0.2, marginTop: 0.2, marginRight: 0.2, marginBottom: 0.2 });
    expect(mixed).toMatchObject({ marginLeft: 0.01, marginTop: 0.02, marginRight: 0.1, marginBottom: 0.05, margin: undefined });
    expect(plain).toMatchObject({ marginLeft: 0.1, marginTop: 0.05, margin: undefined });

    const xml = await slideXml(pptx);
    expect(xml).toContain('lIns="182880"');
    expect(xml).toContain('lIns="9144"');
    expect(xml).toContain('tIns="18288"');

    const reopened = await Presentation.open(await pptx.toBuffer());
    const boxes = reopened.slides.get(0).getTextBoxes();
    expect(boxes.map(box => box.marginLeft)).toEqual([0.2, 0.01, 0.1]);
    expect(boxes.map(box => box.marginTop)).toEqual([0.2, 0.02, 0.05]);
  });

  it('treats re-assigning an inherited default as a no-op', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('defaulted', { name: 'Defaulted' });
    const before = await pptx.toBuffer();
    const opened = await Presentation.open(before);
    const slide = opened.slides.get(0);
    const frame = slide.shapes.findByName('Defaulted')!.textFrame;

    // The document carries no inset attributes, so the value comes from the OOXML defaults.
    expect(await slideXml(opened)).not.toMatch(/Ins="/);
    expect(frame.marginLeft).toBe(0.1);
    expect(slide.dirty).toBe(false);

    // Re-assigning that effective value must neither dirty the slide nor insert lIns="91440".
    frame.marginLeft = 0.1;
    expect(slide.dirty).toBe(false);
    expect(frame.marginLeft).toBe(0.1);
    const xml = await slideXml(opened);
    expect(xml).not.toContain('lIns=');
    expect(await opened.toBuffer()).toEqual(before);

    // Same for the other sides and for the convenience accessor.
    frame.marginTop = 0.05;
    frame.marginRight = 0.1;
    frame.marginBottom = 0.05;
    expect(slide.dirty).toBe(false);
    expect(await opened.toBuffer()).toEqual(before);
    expect(await slideXml(opened)).not.toMatch(/Ins="/);
  });

  it('does not mark the slide dirty when margins are read or unchanged', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('stable', { name: 'Stable', margin: 0.2 });
    const before = await pptx.toBuffer();
    const opened = await Presentation.open(before);
    const frame = opened.slides.get(0).shapes.findByName('Stable')!.textFrame;
    expect(frame.marginLeft).toBe(0.2);
    expect(frame.margin).toBe(0.2);
    frame.marginLeft = 0.2;
    frame.marginTop = 0.2;
    frame.marginRight = 0.2;
    frame.marginBottom = 0.2;
    frame.margin = 0.2;
    expect(await opened.toBuffer()).toEqual(before);
  });

  it('reads malformed or inherited inset attributes as the default', async () => {
    for (const raw of ['constructor', '__proto__', 'toString', 'abc', '', '-1', 'NaN', '1.5e999']) {
      const opened = await withSlideXml(xml => xml.replace(BODY_PR, `<a:bodyPr lIns="${raw}"/>`));
      const frame = opened.slides.get(0).shapes.findByName('External')!.textFrame;
      expect(frame.marginLeft).toBe(0.1);
      expect(frame.margin).toBeUndefined();
    }
  });

  it('rejects invalid sides and values', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('nope', { name: 'Nope' });
    const frame = slide.shapes.findByName('Nope')!.textFrame;

    for (const bad of [NaN, -1, -0.001, Infinity, '0.1', {}, true]) {
      expect(() => { frame.marginLeft = bad as number; }).toThrow(/Invalid text-frame margin/);
      expect(() => { frame.margin = bad as number; }).toThrow(/Invalid text-frame margin/);
      expect(() => slide.addTextBox('x', { margin: bad as number })).toThrow(/Invalid text-frame margin/);
    }
    expect(await slideXml(pptx)).not.toMatch(/Ins="/);
    expect(frame.marginLeft).toBe(0.1);
    frame.marginLeft = null as unknown as undefined; // null means "inherit"
    expect(frame.marginLeft).toBe(0.1);

    expect(() => assertMarginInches(-2)).toThrow(/Invalid text-frame margin/);
    expect(() => assertMarginSide('constructor')).toThrow(/Invalid text-frame margin side/);
    expect(() => assertMarginSide('middle')).toThrow(/Invalid text-frame margin side/);
    expect(() => readMarginInches({}, 'constructor' as TextFrameMarginSide)).toThrow(/Invalid text-frame margin side/);
    expect(() => writeMarginInches({}, '__proto__' as TextFrameMarginSide, 0.1)).toThrow(/Invalid text-frame margin side/);
    expect(() => writeMarginInches({}, 'left', -1)).toThrow(/Invalid text-frame margin/);
  });

  it('resolves margin lookups through own properties only', () => {
    expect(marginAttrFor('left')).toBe('lIns');
    expect(marginAttrFor('top')).toBe('tIns');
    expect(marginAttrFor('right')).toBe('rIns');
    expect(marginAttrFor('bottom')).toBe('bIns');
    for (const inherited of ['constructor', 'toString', 'valueOf', 'hasOwnProperty', '__proto__', 'prototype', 'isPrototypeOf']) {
      expect(marginAttrFor(inherited)).toBeUndefined();
    }
    for (const raw of [undefined, null, 0, 1, {}, [], () => 'left']) expect(marginAttrFor(raw)).toBeUndefined();

    expect(Object.isFrozen(MARGIN_ATTR)).toBe(true);
    expect(Object.isFrozen(MARGIN_DEFAULT_EMU)).toBe(true);
    expect(() => { (MARGIN_ATTR as Record<string, string>).left = 'evil'; }).toThrow();
    expect(marginAttrFor('left')).toBe('lIns');
    expect(MARGIN_DEFAULT_EMU).toEqual({ left: 91440, top: 45720, right: 91440, bottom: 45720 });

    const bodyPr: Record<string, unknown> = {};
    writeMarginInches(bodyPr, 'left', 0.25);
    expect(bodyPr['@_lIns']).toBe('228600');
    expect(readMarginInches(bodyPr, 'left')).toBe(0.25);
    writeMarginInches(bodyPr, 'left', undefined);
    expect(Object.hasOwn(bodyPr, '@_lIns')).toBe(false);
    expect(readMarginInches(bodyPr, 'left')).toBe(0.1);
    expect(readMarginInches(undefined, 'bottom')).toBe(0.05);
  });
});
