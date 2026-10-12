import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Presentation, type Paragraph } from '../src/index.js';

/**
 * Independent acceptance tests for issue #20 (paragraph spacing, indentation and bullets).
 * Written from the acceptance criteria only, following the JSZip pattern of test/alignment.test.ts.
 */

async function slideXml(pptx: Presentation, path = 'ppt/slides/slide1.xml'): Promise<string> {
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  return (await zip.file(path)!.async('string'));
}

function box(text: string, name = 'Body') {
  const pptx = Presentation.create();
  pptx.slides.add().addTextBox(text, { name });
  return { pptx, paragraph: pptx.slides.get(0).shapes.findByName(name)!.textFrame.paragraphs[0] };
}

function paragraphsOf(pptx: Presentation, name = 'Body'): Paragraph[] {
  return pptx.slides.get(0).shapes.findByName(name)!.textFrame.paragraphs;
}

/** Injects a pre-existing a:pPr into the first paragraph of slide1, like alignment.test.ts does. */
async function withInjectedPPr(replacement: string, name = 'Body') {
  const pptx = Presentation.create();
  pptx.slides.add().addTextBox('target\nsecond', { name });
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  const xml = (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace('<a:pPr/>', replacement);
  zip.file('ppt/slides/slide1.xml', xml);
  const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
  return { pptx: opened, paragraph: opened.slides.get(0).shapes.findByName(name)!.textFrame.paragraphs[0] };
}

describe('#20 acceptance: round-trip write/read of every property', () => {
  it('round-trips level 0 and 8 through save+reopen and writes lvl', async () => {
    for (const level of [0, 8] as const) {
      const { pptx, paragraph } = box('Outline');
      expect(paragraph.level).toBeUndefined();
      paragraph.level = level;
      expect(paragraph.level).toBe(level);
      expect(await slideXml(pptx)).toContain(`lvl="${level}"`);
      const reopened = await Presentation.open(await pptx.toBuffer());
      expect(reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0].level).toBe(level);
    }
  });

  it('round-trips leftMargin and indentation as EMU, including negative values', async () => {
    const { pptx, paragraph } = box('Indented');
    paragraph.leftMargin = -274320;
    paragraph.indentation = 274320;
    const xml = await slideXml(pptx);
    expect(xml).toMatch(/marL="-274320"/);
    expect(xml).toMatch(/indent="274320"/);
    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect(back.leftMargin).toBe(-274320);
    expect(back.indentation).toBe(274320);
  });

  it('round-trips spaceBefore and spaceAfter in points, including 0', async () => {
    const { pptx, paragraph } = box('Spaced');
    paragraph.spaceBefore = 18;
    paragraph.spaceAfter = 0;
    const xml = await slideXml(pptx);
    expect(xml).toContain('<a:spcBef><a:spcPts val="1800"/></a:spcBef>');
    expect(xml).toContain('<a:spcAft><a:spcPts val="0"/></a:spcAft>');
    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect(back.spaceBefore).toBe(18);
    expect(back.spaceAfter).toBe(0);
  });

  it('round-trips bullet none, bullet char and bullet autoNumber', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('a\nb\nc\nd', { name: 'List' });
    const [none, charBullet, auto, plain] = paragraphsOf(pptx, 'List');
    expect(plain.bullet).toBeUndefined();

    none.bullet = 'none';
    charBullet.bullet = { type: 'char', char: '•' };
    auto.bullet = { type: 'autoNumber', autoNumber: 'arabicDash' };
    expect(none.bullet).toEqual({ type: 'none' });
    expect(charBullet.bullet).toEqual({ type: 'char', char: '•' });
    expect(auto.bullet).toEqual({ type: 'autoNumber', autoNumber: 'arabicDash' });

    const xml = await slideXml(pptx);
    expect(xml).toContain('<a:buNone/>');
    expect(xml).toContain('<a:buChar char="•"/>');
    expect(xml).toContain('<a:buAutoNum type="arabicDash"/>');

    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('List')!.textFrame.paragraphs;
    expect(back[0].bullet).toEqual({ type: 'none' });
    expect(back[1].bullet).toEqual({ type: 'char', char: '•' });
    expect(back[2].bullet).toEqual({ type: 'autoNumber', autoNumber: 'arabicDash' });
    expect(back[3].bullet).toBeUndefined();
  });
});

describe('#20 acceptance: undefined means inherited', () => {
  it('emits nothing for undefined and removes the attribute/node when reset to undefined', async () => {
    const { pptx, paragraph } = box('Clear');
    paragraph.level = 3;
    paragraph.leftMargin = 457200;
    paragraph.indentation = -228600;
    paragraph.spaceBefore = 12;
    paragraph.spaceAfter = 6;
    paragraph.bullet = { type: 'char', char: '‣' };
    const set = await slideXml(pptx);
    for (const fragment of ['lvl="3"', 'marL="457200"', 'indent="-228600"', '<a:spcBef>', '<a:spcAft>', '<a:buChar']) {
      expect(set).toContain(fragment);
    }

    paragraph.level = undefined;
    paragraph.leftMargin = undefined;
    paragraph.indentation = undefined;
    paragraph.spaceBefore = undefined;
    paragraph.spaceAfter = undefined;
    paragraph.bullet = undefined;
    expect(await slideXml(pptx)).toBe(await slideXml(box('Clear').pptx));
    for (const value of [paragraph.level, paragraph.leftMargin, paragraph.indentation, paragraph.spaceBefore, paragraph.spaceAfter, paragraph.bullet]) {
      expect(value).toBeUndefined();
    }

    // And after a real save+reopen cycle.
    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect([back.level, back.leftMargin, back.indentation, back.spaceBefore, back.spaceAfter, back.bullet]).toEqual([
      undefined, undefined, undefined, undefined, undefined, undefined,
    ]);
    expect(await slideXml(reopened)).not.toMatch(/lvl=|marL=|indent=|a:spc|a:bu(None|Char|AutoNum)/);
  });
});

describe('#20 acceptance: invalid values raise TypeError listing valid options', () => {
  const BAD: unknown[] = [-1, 9, 1.5, NaN, Infinity, 'x', '457200', '', {}, [], true, null as unknown, 'constructor', '__proto__', 'toString'];

  it('rejects invalid level values with a descriptive TypeError', () => {
    const { paragraph } = box('Level');
    for (const bad of BAD.filter(v => v !== null)) {
      expect(() => { paragraph.level = bad as number; }, String(bad)).toThrow(TypeError);
      expect(() => { paragraph.level = bad as number; }).toThrow(/Invalid paragraph level/);
      expect(() => { paragraph.level = bad as number; }).toThrow(/0 and 8/);
    }
    expect(paragraph.level).toBeUndefined();
    paragraph.level = null as unknown as number;
    expect(paragraph.level).toBeUndefined();
  });

  it('rejects invalid marL/indent EMU values', () => {
    const { paragraph } = box('Emu');
    for (const bad of [-1.5, NaN, Infinity, 'x', '12pt', {}, [], 'constructor', '__proto__']) {
      expect(() => { paragraph.leftMargin = bad as number; }, String(bad)).toThrow(/Invalid paragraph leftMargin/);
      expect(() => { paragraph.indentation = bad as number; }, String(bad)).toThrow(/Invalid paragraph indentation/);
    }
    expect(paragraph.leftMargin).toBeUndefined();
    expect(paragraph.indentation).toBeUndefined();
  });

  it('rejects invalid spacing values', () => {
    const { paragraph } = box('Spacing');
    for (const bad of [NaN, Infinity, 'x', '12', {}, 'constructor', '__proto__']) {
      expect(() => { paragraph.spaceBefore = bad as number; }, String(bad)).toThrow(/Invalid paragraph spaceBefore/);
      expect(() => { paragraph.spaceAfter = bad as number; }, String(bad)).toThrow(/Invalid paragraph spaceAfter/);
    }
    expect(paragraph.spaceBefore).toBeUndefined();
    expect(paragraph.spaceAfter).toBeUndefined();
  });

  it('rejects invalid bullets, including prototype-inherited names', () => {
    const { paragraph } = box('Bullet', 'List');
    for (const bad of ['disc', '', 'constructor', '__proto__', 'toString', 1, {}, [], { type: 'char', char: '' }, { type: 'char', char: 7 }, { type: 'autoNumber', autoNumber: 'roman' }, { type: 'autoNumber', autoNumber: 'constructor' }, { type: 'autoNumber', number: 0 }, { type: 'none', number: -2 }]) {
      expect(() => { paragraph.bullet = bad as never; }, String(bad)).toThrow(TypeError);
      expect(() => { paragraph.bullet = bad as never; }).toThrow(/Invalid paragraph bullet|Invalid bullet/);
    }
    expect(paragraph.bullet).toBeUndefined();
    // The message enumerates the supported autoNumber options.
    try {
      paragraph.bullet = { type: 'autoNumber', autoNumber: 'roman' as never };
      expect.unreachable();
    } catch (error) {
      expect((error as Error).message).toContain('arabicPeriod');
      expect((error as Error).message).toContain('circleNumDbChsFullWidth');
    }
  });
});

describe('#20 acceptance: non-destructive edits on an existing a:pPr', () => {
  const EXISTING =
    '<a:pPr marL="457200" indent="-228600" algn="ctr" rtl="1">' +
    '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>' +
    '<a:buFont typeface="Arial"/><a:buChar char="•"/>' +
    '<a:defRPr sz="2400"/></a:pPr>';

  it('only changes the targeted nodes and keeps the schema child order', async () => {
    const { pptx, paragraph } = await withInjectedPPr(EXISTING);
    expect(paragraph.leftMargin).toBe(457200);
    expect(paragraph.indentation).toBe(-228600);
    expect(paragraph.alignment).toBe('center');
    expect(paragraph.bullet).toEqual({ type: 'char', char: '•' });

    paragraph.leftMargin = -274320;
    paragraph.spaceBefore = 18;
    paragraph.bullet = 'none';
    const xml = await slideXml(pptx);
    expect(xml).toMatch(
      /<a:pPr marL="-274320" indent="-228600" algn="ctr" rtl="1"><a:lnSpc><a:spcPct val="90000"\/><\/a:lnSpc><a:spcBef><a:spcPts val="1800"\/><\/a:spcBef><a:buFont typeface="Arial"\/><a:buNone\/><a:defRPr sz="2400"\/><\/a:pPr>/,
    );
    expect(xml).not.toContain('a:buChar');
    // Second paragraph untouched.
    expect(xml.match(/<a:pPr\/>/g)?.length).toBe(1);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect(back.leftMargin).toBe(-274320);
    expect(back.indentation).toBe(-228600);
    expect(back.alignment).toBe('center');
    expect(back.spaceBefore).toBe(18);
    expect(back.bullet).toEqual({ type: 'none' });
    expect(back.text).toBe('target');
  });

  it('creates a:pPr as the first child of a:p', async () => {
    const { pptx, paragraph } = box('Fresh');
    paragraph.spaceAfter = 6;
    paragraph.level = 8;
    expect(await slideXml(pptx)).toMatch(/<a:p><a:pPr lvl="8"><a:spcAft><a:spcPts val="600"\/><\/a:spcAft><\/a:pPr><a:r>/);
  });
});

describe('#20 acceptance: idempotency and independence', () => {
  it('setting the same value again does not duplicate or reorder children', async () => {
    const { pptx, paragraph } = box('Idem');
    paragraph.level = 2;
    paragraph.leftMargin = 914400;
    paragraph.spaceBefore = 12;
    paragraph.spaceAfter = 4;
    paragraph.bullet = { type: 'autoNumber', autoNumber: 'alphaLcParen' };
    const first = await slideXml(pptx);
    for (let i = 0; i < 3; i += 1) {
      paragraph.level = 2;
      paragraph.leftMargin = 914400;
      paragraph.spaceBefore = 12;
      paragraph.spaceAfter = 4;
      paragraph.bullet = { type: 'autoNumber', autoNumber: 'alphaLcParen' };
    }
    const again = await slideXml(pptx);
    expect(again).toBe(first);
    for (const fragment of ['<a:pPr ', '<a:spcBef>', '<a:spcAft>', '<a:buAutoNum']) {
      expect((again.split(fragment).length - 1) >= 1).toBe(true);
    }
    expect((again.match(/<a:spcBef>/g) ?? []).length).toBe(1);
    expect((again.match(/<a:spcAft>/g) ?? []).length).toBe(1);
    expect((again.match(/<a:buAutoNum/g) ?? []).length).toBe(1);
    expect((again.match(/<a:pPr /g) ?? []).length).toBe(1);
    expect((again.match(/lvl="2"/g) ?? []).length).toBe(1);
  });

  it('keeps paragraphs of the same text frame independent', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('one\ntwo\nthree', { name: 'Frame' });
    const [one, two, three] = paragraphsOf(pptx, 'Frame');
    one.level = 4;
    one.spaceBefore = 9;
    two.bullet = 'none';
    expect(two.level).toBeUndefined();
    expect(two.spaceBefore).toBeUndefined();
    expect(three.level).toBeUndefined();
    expect(three.bullet).toBeUndefined();
    expect(one.spaceAfter).toBeUndefined();
    const xml = await slideXml(pptx);
    expect((xml.match(/lvl="4"/g) ?? []).length).toBe(1);
    expect((xml.match(/<a:buNone\/>/g) ?? []).length).toBe(1);
    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('Frame')!.textFrame.paragraphs;
    expect(back.map(p => p.level)).toEqual([4, undefined, undefined]);
    expect(back.map(p => p.spaceBefore)).toEqual([9, undefined, undefined]);
    expect(back[1].bullet).toEqual({ type: 'none' });
    expect(back[0].bullet).toBeUndefined();
  });

  it('keeps slides independent', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('first', { name: 'Body' });
    pptx.slides.add().addTextBox('second', { name: 'Body' });
    const before2 = await slideXml(pptx, 'ppt/slides/slide2.xml');
    const first = pptx.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    first.level = 5;
    first.leftMargin = 457200;
    first.spaceAfter = 3;
    first.bullet = { type: 'char', char: '•' };
    expect(await slideXml(pptx, 'ppt/slides/slide2.xml')).toBe(before2);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const back1 = reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    const back2 = reopened.slides.get(1).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect([back1.level, back1.leftMargin, back1.spaceAfter, back1.bullet]).toEqual([5, 457200, 3, { type: 'char', char: '•' }]);
    expect([back2.level, back2.leftMargin, back2.spaceAfter, back2.bullet]).toEqual([undefined, undefined, undefined, undefined]);
    expect(await slideXml(reopened, 'ppt/slides/slide2.xml')).toBe(before2);
  });
});
