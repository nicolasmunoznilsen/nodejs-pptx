import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  assertMarginInches,
  assertMarginSide,
  marginAttrFor,
  readMarginInches,
  writeMarginInches,
  type TextFrameMarginSide,
  type TextBoxInfo,
} from '../src/index.js';

/** Raw XML of slide 1, straight out of the package (no library round-trip involved). */
async function slideXml(source: Presentation | Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(source instanceof Presentation ? await source.toBuffer() : source);
  const file = zip.file('ppt/slides/slide1.xml');
  expect(file, 'ppt/slides/slide1.xml must exist').not.toBeNull();
  return (file as JSZip.JSZipObject).async('string');
}

/** Every stored part of the package, keyed by zip path. */
async function parts(source: Presentation | Buffer): Promise<Map<string, Buffer>> {
  const zip = await JSZip.loadAsync(source instanceof Presentation ? await source.toBuffer() : source);
  const result = new Map<string, Buffer>();
  await Promise.all(Object.entries(zip.files).filter(([, entry]) => !entry.dir).map(async ([name, entry]) => {
    result.set(name, await entry.async('nodebuffer'));
  }));
  return result;
}

/** A saved deck with one text box carrying one explicit inset per side. */
async function deckWithInsets(): Promise<{ buffer: Buffer; box: TextBoxInfo }> {
  const pptx = Presentation.create();
  const slide = pptx.slides.add();
  const box = slide.addTextBox('margins', {
    name: 'Box',
    marginLeft: 0.5,
    marginTop: 0.4,
    marginRight: 0.3,
    marginBottom: 0.2,
  });
  return { buffer: await pptx.toBuffer(), box };
}

const SIDES: Array<{ side: TextFrameMarginSide; attr: string; key: keyof TextBoxInfo; emu: string; inches: number; defaultInches: number }> = [
  { side: 'left', attr: 'lIns', key: 'marginLeft', emu: '457200', inches: 0.5, defaultInches: 0.1 },
  { side: 'top', attr: 'tIns', key: 'marginTop', emu: '365760', inches: 0.4, defaultInches: 0.05 },
  { side: 'right', attr: 'rIns', key: 'marginRight', emu: '274320', inches: 0.3, defaultInches: 0.1 },
  { side: 'bottom', attr: 'bIns', key: 'marginBottom', emu: '182880', inches: 0.2, defaultInches: 0.05 },
];

const SETTER: Record<TextFrameMarginSide, 'marginLeft' | 'marginTop' | 'marginRight' | 'marginBottom'> = {
  left: 'marginLeft',
  top: 'marginTop',
  right: 'marginRight',
  bottom: 'marginBottom',
};

describe('text-frame margins (independent acceptance)', () => {
  // (a) defaults
  it('(a) reports 0.1/0.05/0.1/0.05 and stores no *Ins attributes on a fresh text box', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('fresh', { name: 'Fresh' });

    const live = slide.shapes.findByName('Fresh')!.textFrame;
    expect(live.marginLeft).toBe(0.1);
    expect(live.marginTop).toBe(0.05);
    expect(live.marginRight).toBe(0.1);
    expect(live.marginBottom).toBe(0.05);

    const buffer = await pptx.toBuffer();
    const xml = await slideXml(buffer);
    expect(xml).not.toMatch(/(lIns|tIns|rIns|bIns)=/);

    const reopened = await Presentation.open(buffer);
    const frame = reopened.slides.get(0).shapes.findByName('Fresh')!.textFrame;
    expect([frame.marginLeft, frame.marginTop, frame.marginRight, frame.marginBottom]).toEqual([0.1, 0.05, 0.1, 0.05]);
    expect(reopened.slides.get(0).getTextBoxes()[0]).toMatchObject({
      marginLeft: 0.1, marginTop: 0.05, marginRight: 0.1, marginBottom: 0.05,
    });
    // Library-level helpers agree with the accessors.
    expect(readMarginInches(undefined, 'left')).toBe(0.1);
    expect(readMarginInches(undefined, 'top')).toBe(0.05);
    expect(readMarginInches(undefined, 'right')).toBe(0.1);
    expect(readMarginInches(undefined, 'bottom')).toBe(0.05);
  });

  // (b) marginLeft = 0.25
  it('(b) writes lIns="228600" for marginLeft = 0.25 and reopens as 0.25', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('left', { name: 'Left' });
    const frame = slide.shapes.findByName('Left')!.textFrame;

    frame.marginLeft = 0.25;
    expect(frame.marginLeft).toBe(0.25);
    const xml = await slideXml(pptx);
    expect(xml).toContain('lIns="228600"');
    expect(xml).not.toMatch(/(tIns|rIns|bIns)=/);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const again = reopened.slides.get(0).shapes.findByName('Left')!.textFrame;
    expect(again.marginLeft).toBe(0.25);
    expect(await slideXml(await reopened.toBuffer())).toContain('lIns="228600"');
    // Second generation round-trip keeps the value.
    const second = await Presentation.open(await reopened.toBuffer());
    expect(second.slides.get(0).shapes.findByName('Left')!.textFrame.marginLeft).toBe(0.25);
  });

  // (c) clearing one side
  it.each(SIDES)('(c) %s = undefined removes only that attribute and reverts to its default', async entry => {
    const { buffer } = await deckWithInsets();
    const before = await slideXml(buffer);
    for (const other of SIDES) expect(before).toContain(`${other.attr}="${other.emu}"`);

    const opened = await Presentation.open(buffer);
    const slide = opened.slides.get(0);
    const frame = slide.shapes.findByName('Box')!.textFrame;
    expect(frame[SETTER[entry.side]]).toBe(entry.inches);

    frame[SETTER[entry.side]] = undefined;
    expect(frame[SETTER[entry.side]]).toBe(entry.defaultInches);

    const xml = await slideXml(opened);
    expect(xml).not.toContain(`${entry.attr}=`);
    for (const other of SIDES.filter(other => other.side !== entry.side)) {
      expect(xml).toContain(`${other.attr}="${other.emu}"`);
      expect(frame[SETTER[other.side]]).toBe(other.inches);
    }

    const reopened = await Presentation.open(await opened.toBuffer());
    const again = reopened.slides.get(0).shapes.findByName('Box')!.textFrame;
    expect(again[SETTER[entry.side]]).toBe(entry.defaultInches);
    for (const other of SIDES.filter(other => other.side !== entry.side)) {
      expect(again[SETTER[other.side]]).toBe(other.inches);
    }
  });

  it('(c) addTextBox options are the only source of explicit insets and can be cleared the same way', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('one', { name: 'One', marginLeft: 0.1 });
    const frame = slide.shapes.findByName('One')!.textFrame;
    // An explicit attribute equal to the default is still explicit.
    expect(await slideXml(pptx)).toContain('lIns="91440"');
    frame.marginLeft = undefined;
    expect(frame.marginLeft).toBe(0.1);
    expect(await slideXml(pptx)).not.toContain('lIns=');
  });

  // (d) unrelated markup
  it('(d) leaves unrelated bodyPr attributes, children and all other slide parts intact', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('other', { name: 'Other', fontSize: 18, color: 'FF0000', alignment: 'center' });
    pptx.slides.add().addTextBox('second slide', { name: 'Second' });
    const before = await parts(pptx);

    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const path = 'ppt/slides/slide1.xml';
    let xml = (await zip.file(path)!.async('string'));
    expect(xml).toContain('<a:bodyPr wrap="square"');
    xml = xml.replace(
      /<a:bodyPr[^>]*\/>/,
      '<a:bodyPr wrap="square" anchor="ctr" rtlCol="0" lIns="914400"><a:normAutofit/></a:bodyPr>',
    );
    zip.file(path, xml);
    const injected = await zip.generateAsync({ type: 'nodebuffer' });
    expect(await slideXml(injected)).toContain('<a:normAutofit/>');

    const opened = await Presentation.open(injected);
    const slide = opened.slides.get(0);
    const frame = slide.shapes.findByName('Other')!.textFrame;
    expect(frame.marginLeft).toBe(1);
    expect(frame.marginTop).toBe(0.05);

    frame.marginTop = 0.2;
    frame.marginRight = undefined;

    const after = await parts(opened);
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
    for (const [name, bytes] of after) {
      if (name === path) continue;
      expect(bytes.equals(before.get(name)!), `${name} must be byte-identical`).toBe(true);
    }

    const edited = await slideXml(opened);
    expect(edited).toContain('wrap="square"');
    expect(edited).toContain('anchor="ctr"');
    expect(edited).toContain('rtlCol="0"');
    expect(edited).toContain('lIns="914400"');
    expect(edited).toContain('tIns="182880"');
    expect(edited).not.toContain('rIns=');
    expect(edited).toContain('<a:normAutofit/>');
    expect(edited).not.toContain('<a:normAutofit/><a:normAutofit');
    expect(await opened.slides.get(0).text).toBe('other');
    expect(await opened.slides.get(1).text).toBe('second slide');

    const reopened = await Presentation.open(await opened.toBuffer());
    expect(reopened.slides.length).toBe(2);
    const again = reopened.slides.get(0).shapes.findByName('Other')!.textFrame;
    expect(again.marginLeft).toBe(1);
    expect(again.marginTop).toBe(0.2);
    expect(again.marginRight).toBe(0.1);
    expect(await reopened.slides.get(1).text).toBe('second slide');
  });

  // (e) read-only + no-op writes
  it('(e) reading margins leaves the file byte-identical and re-assigning the effective value is a no-op', async () => {
    const explicit = Presentation.create();
    explicit.slides.add().addTextBox('stable', { name: 'Stable', marginLeft: 0.25, marginRight: 0.25 });
    const beforeExplicit = await explicit.toBuffer();
    const openedExplicit = await Presentation.open(beforeExplicit);
    const slideExplicit = openedExplicit.slides.get(0);
    const frameExplicit = slideExplicit.shapes.findByName('Stable')!.textFrame;

    expect(frameExplicit.marginLeft).toBe(0.25);
    expect(frameExplicit.marginTop).toBe(0.05);
    expect(frameExplicit.marginRight).toBe(0.25);
    expect(frameExplicit.marginBottom).toBe(0.05);
    expect(slideExplicit.dirty).toBe(false);
    expect(await openedExplicit.toBuffer()).toEqual(beforeExplicit);

    // Re-assigning the current effective value must not dirty the slide.
    frameExplicit.marginLeft = 0.25;
    frameExplicit.marginRight = 0.25;
    expect(slideExplicit.dirty).toBe(false);
    expect(await openedExplicit.toBuffer()).toEqual(beforeExplicit);

    // Inherited defaults are "current effective values" too.
    const defaulted = Presentation.create();
    defaulted.slides.add().addTextBox('defaulted', { name: 'Defaulted' });
    const beforeDefaulted = await defaulted.toBuffer();
    const openedDefaulted = await Presentation.open(beforeDefaulted);
    const slideDefaulted = openedDefaulted.slides.get(0);
    const frameDefaulted = slideDefaulted.shapes.findByName('Defaulted')!.textFrame;
    expect(await slideXml(openedDefaulted)).not.toMatch(/(lIns|tIns|rIns|bIns)=/);

    frameDefaulted.marginLeft = 0.1;
    frameDefaulted.marginTop = 0.05;
    frameDefaulted.marginRight = 0.1;
    frameDefaulted.marginBottom = 0.05;
    expect(slideDefaulted.dirty).toBe(false);
    expect(await slideXml(openedDefaulted)).not.toMatch(/(lIns|tIns|rIns|bIns)=/);
    expect(await openedDefaulted.toBuffer()).toEqual(beforeDefaulted);

    // A real edit does dirty the slide, proving the assertions above are meaningful.
    frameDefaulted.marginTop = 0.3;
    expect(slideDefaulted.dirty).toBe(true);
    expect(await openedDefaulted.toBuffer()).not.toEqual(beforeDefaulted);
  });

  // (f) foreign-authored insets
  it('(f) reads insets written by an external tool, including lIns="0"', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('external', { name: 'External' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const path = 'ppt/slides/slide1.xml';
    const xml = (await zip.file(path)!.async('string')).replace(
      /<a:bodyPr[^>]*\/>/,
      '<a:bodyPr lIns="0" rIns="182880"><a:normAutofit/></a:bodyPr>',
    );
    zip.file(path, xml);
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    const injected = await slideXml(buffer);
    expect(injected).toContain('<a:bodyPr lIns="0" rIns="182880">');
    expect(injected).toContain('<a:normAutofit/>');

    const opened = await Presentation.open(buffer);
    const slide = opened.slides.get(0);
    const frame = slide.shapes.findByName('External')!.textFrame;
    expect(frame.marginLeft).toBe(0);
    expect(frame.marginRight).toBe(0.2);
    expect(frame.marginTop).toBe(0.05);
    expect(frame.marginBottom).toBe(0.05);
    expect(frame.margin).toBeUndefined();
    expect(slide.dirty).toBe(false);
    // Opening and reading must not rewrite the foreign XML.
    expect(await slideXml(opened)).toBe(await slideXml(buffer));

    // Editing one side keeps the untouched foreign attributes.
    frame.marginBottom = 0.4;
    const edited = await slideXml(opened);
    expect(edited).toContain('lIns="0"');
    expect(edited).toContain('rIns="182880"');
    expect(edited).toContain('bIns="365760"');
    expect(edited).toContain('<a:normAutofit/>');
    const reopened = await Presentation.open(await opened.toBuffer());
    const again = reopened.slides.get(0).shapes.findByName('External')!.textFrame;
    expect([again.marginLeft, again.marginRight, again.marginBottom]).toEqual([0, 0.2, 0.4]);

    // An explicit 0 is a real value: clearing it reverts to the 0.1" default.
    const cleared = await Presentation.open(buffer);
    const clearedFrame = cleared.slides.get(0).shapes.findByName('External')!.textFrame;
    expect(clearedFrame.marginLeft).toBe(0);
    clearedFrame.marginLeft = undefined;
    expect(clearedFrame.marginLeft).toBe(0.1);
    expect(await slideXml(cleared)).not.toContain('lIns=');
    expect(await slideXml(cleared)).toContain('rIns="182880"');
  });

  // (g) invalid inputs
  it('(g) rejects invalid sides and values with TypeError', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('invalid', { name: 'Invalid' });
    const frame = slide.shapes.findByName('Invalid')!.textFrame;
    const before = await pptx.toBuffer();

    // Bad sides.
    for (const bad of ['middle', 'constructor', '__proto__', 'toString', 'valueOf', 'LEFT', '']) {
      expect(
        () => readMarginInches({}, bad as TextFrameMarginSide),
        `readMarginInches(${bad})`,
      ).toThrow(TypeError);
      expect(
        () => writeMarginInches({}, bad as TextFrameMarginSide, 0.1),
        `writeMarginInches(${bad})`,
      ).toThrow(TypeError);
      expect(() => assertMarginSide(bad)).toThrow(TypeError);
      expect(marginAttrFor(bad)).toBeUndefined();
    }
    // Prototype names must not leak an inherited value into generated XML.
    const polluted: Record<string, unknown> = {};
    expect(() => writeMarginInches(polluted, '__proto__' as TextFrameMarginSide, 0.1)).toThrow(TypeError);
    expect(Object.keys(polluted)).toEqual([]);

    // Bad values.
    for (const bad of [NaN, -1, -0.0001, -Infinity, Infinity]) {
      expect(() => { frame.marginLeft = bad; }).toThrow(TypeError);
      expect(() => { frame.marginTop = bad; }).toThrow(TypeError);
      expect(() => { frame.margin = bad; }).toThrow(TypeError);
      expect(() => assertMarginInches(bad)).toThrow(TypeError);
      expect(() => slide.addTextBox('x', { marginLeft: bad })).toThrow(TypeError);
      expect(() => slide.addTextBox('x', { margin: bad })).toThrow(TypeError);
      expect(() => writeMarginInches({}, 'left', bad)).toThrow(TypeError);
    }

    // Nothing was written: the serialized slide is byte-identical to the pre-attempt snapshot.
    expect(await pptx.toBuffer()).toEqual(before);
    expect(await slideXml(pptx)).not.toMatch(/(lIns|tIns|rIns|bIns)=/);
    expect(frame.marginLeft).toBe(0.1);
  });

  // (h) the convenience accessor
  it('(h) margin reads the shared value only when all four agree and writes all four', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('uniform', { name: 'Uniform' });
    const frame = slide.shapes.findByName('Uniform')!.textFrame;

    // Defaults disagree (0.1 vs 0.05), so the aggregate reports undefined.
    expect(frame.margin).toBeUndefined();

    frame.margin = 0.4;
    expect(frame.margin).toBe(0.4);
    expect([frame.marginLeft, frame.marginTop, frame.marginRight, frame.marginBottom]).toEqual([0.4, 0.4, 0.4, 0.4]);
    const xml = await slideXml(pptx);
    for (const attr of ['lIns', 'tIns', 'rIns', 'bIns']) expect(xml).toContain(`${attr}="365760"`);

    // One side differs -> aggregate is undefined again, and it stays undefined after reopen.
    frame.marginRight = 0.05;
    expect(frame.margin).toBeUndefined();
    const mixed = await Presentation.open(await pptx.toBuffer());
    expect(mixed.slides.get(0).shapes.findByName('Uniform')!.textFrame.margin).toBeUndefined();

    // Re-uniforming writes the remaining sides only.
    frame.margin = 0.4;
    expect(frame.margin).toBe(0.4);
    const uniformXml = await slideXml(pptx);
    for (const attr of ['lIns', 'tIns', 'rIns', 'bIns']) expect(uniformXml).toContain(`${attr}="365760"`);

    // Clearing through the accessor removes all four attributes.
    frame.margin = undefined;
    expect(await slideXml(pptx)).not.toMatch(/(lIns|tIns|rIns|bIns)=/);
    expect(frame.margin).toBeUndefined();

    // Equal-by-default frames report the shared value.
    const equal = Presentation.create();
    equal.slides.add().addTextBox('four', { name: 'Four', marginLeft: 0.3, marginTop: 0.3, marginRight: 0.3, marginBottom: 0.3 });
    const equalFrame = equal.slides.get(0).shapes.findByName('Four')!.textFrame;
    expect(equalFrame.margin).toBe(0.3);
    const reopenedEqual = await Presentation.open(await equal.toBuffer());
    expect(reopenedEqual.slides.get(0).shapes.findByName('Four')!.textFrame.margin).toBe(0.3);
  });

  // (i) margin: 0
  it('(i) addTextBox({ margin: 0 }) round-trips to 0 on all four sides', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const box = slide.addTextBox('zero', { name: 'Zero', margin: 0 });
    expect(box).toMatchObject({ margin: 0, marginLeft: 0, marginTop: 0, marginRight: 0, marginBottom: 0 });

    const xml = await slideXml(pptx);
    for (const attr of ['lIns', 'tIns', 'rIns', 'bIns']) expect(xml).toContain(`${attr}="0"`);

    const buffer = await pptx.toBuffer();
    const reopened = await Presentation.open(buffer);
    const frame = reopened.slides.get(0).shapes.findByName('Zero')!.textFrame;
    expect(frame.marginLeft).toBe(0);
    expect(frame.marginTop).toBe(0);
    expect(frame.marginRight).toBe(0);
    expect(frame.marginBottom).toBe(0);
    expect(frame.margin).toBe(0);
    expect(reopened.slides.get(0).getTextBoxes()[0]).toMatchObject({
      margin: 0, marginLeft: 0, marginTop: 0, marginRight: 0, marginBottom: 0,
    });
    // 0 must not be confused with "unset": reopening is still not dirty and stays byte-identical.
    expect(reopened.slides.get(0).dirty).toBe(false);
    expect(await reopened.toBuffer()).toEqual(buffer);
    const second = await Presentation.open(await reopened.toBuffer());
    expect(second.slides.get(0).getTextBoxes()[0].margin).toBe(0);
  });
});
