import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  MAX_PARAGRAPH_LEVEL,
  MIN_PARAGRAPH_LEVEL,
  Paragraph,
  Presentation,
  centipointsToPoints,
  emuToPoints,
  orderParagraphProperties,
  pointsToCentipoints,
  pointsToEmu,
  type ParagraphBullet,
} from '../src/index.js';
import { readBullet, writeBullet } from '../src/text/paragraph-format.js';
import type { XmlNode } from '../src/ooxml/xml.js';

async function slideXml(pptx: Presentation): Promise<string> {
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  return (await zip.file('ppt/slides/slide1.xml')!.async('string'));
}

function paragraphOf(text: string, name = 'Body'): Paragraph {
  const pptx = Presentation.create();
  pptx.slides.add().addTextBox(text, { name });
  return pptx.slides.get(0).shapes.findByName(name)!.textFrame.paragraphs[0];
}

describe('paragraph format helpers', () => {
  it('exposes unit conversions', () => {
    expect(pointsToCentipoints(6)).toBe(600);
    expect(centipointsToPoints(600)).toBe(6);
    expect(pointsToEmu(1)).toBe(12700);
    expect(emuToPoints(12700)).toBe(1);
    expect(emuToPoints(pointsToEmu(12))).toBe(12);
    expect([MIN_PARAGRAPH_LEVEL, MAX_PARAGRAPH_LEVEL]).toEqual([0, 8]);
  });

  it('reorders a:pPr children without losing attributes or unknown children', () => {
    const pPr: XmlNode = {
      '@_marL': '457200',
      '@_rtl': '1',
      'a:defRPr': '',
      'a:buChar': { '@_char': '•' },
      'a:spcAft': { 'a:spcPts': { '@_val': '600' } },
      'a:lnSpc': { 'a:spcPct': { '@_val': '100000' } },
      'a:tabLst': { 'a:tab': { '@_pos': '914400', '@_algn': 'l' } },
      'a:futureElement': { '@_keep': 'yes' },
    };
    orderParagraphProperties(pPr);
    expect(Object.keys(pPr)).toEqual([
      '@_marL',
      '@_rtl',
      'a:lnSpc',
      'a:spcAft',
      'a:buChar',
      'a:tabLst',
      'a:defRPr',
      'a:futureElement',
    ]);
    expect(pPr['a:tabLst']).toEqual({ 'a:tab': { '@_pos': '914400', '@_algn': 'l' } });
  });
});

describe('paragraph level', () => {
  it('writes, reads, clears, and round-trips lvl', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Outline', { name: 'Body' });
    const shape = pptx.slides.get(0).shapes.findByName('Body')!;
    const paragraph = shape.textFrame.paragraphs[0];

    expect(paragraph.level).toBeUndefined();
    paragraph.level = 2;
    expect(paragraph.level).toBe(2);
    expect(await slideXml(pptx)).toContain('<a:pPr lvl="2"');

    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0].level).toBe(2);

    paragraph.level = undefined;
    expect(paragraph.level).toBeUndefined();
    expect(await slideXml(pptx)).not.toContain('lvl=');
  });

  it('accepts both boundary levels and rejects everything else', () => {
    const paragraph = paragraphOf('Outline');
    paragraph.level = 0;
    expect(paragraph.level).toBe(0);
    paragraph.level = 8;
    expect(paragraph.level).toBe(8);
    for (const bad of [-1, 9, 1.5, '2', {}, NaN, Infinity, true]) {
      expect(() => { paragraph.level = bad as number; }).toThrow(TypeError);
      expect(() => { paragraph.level = bad as number; }).toThrow(/Invalid paragraph level/);
    }
    expect(paragraph.level).toBe(8);
  });
});

describe('paragraph indentation and margins', () => {
  it('writes marL and indent as integer EMU and round-trips them', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Indented', { name: 'Body' });
    const paragraph = pptx.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];

    expect(paragraph.leftMargin).toBeUndefined();
    expect(paragraph.indentation).toBeUndefined();

    paragraph.leftMargin = 457200;
    paragraph.indentation = -228600;
    expect(paragraph.leftMargin).toBe(457200);
    expect(paragraph.indentation).toBe(-228600);
    expect(await slideXml(pptx)).toMatch(/<a:pPr marL="457200" indent="-228600"/);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect(back.leftMargin).toBe(457200);
    expect(back.indentation).toBe(-228600);

    back.leftMargin = undefined;
    expect(back.leftMargin).toBeUndefined();
  });

  it('rejects non-integer EMU values', () => {
    const paragraph = paragraphOf('Indented');
    for (const bad of [1.5, '457200', {}, NaN, Infinity]) {
      expect(() => { paragraph.leftMargin = bad as number; }).toThrow(/Invalid paragraph leftMargin/);
      expect(() => { paragraph.indentation = bad as number; }).toThrow(/Invalid paragraph indentation/);
    }
    expect(paragraph.leftMargin).toBeUndefined();
  });
});

describe('paragraph spacing', () => {
  it('writes spcBef and spcAft in centipoints, schema-ordered', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Spaced', { name: 'Body' });
    const paragraph = pptx.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];

    expect(paragraph.spaceBefore).toBeUndefined();
    paragraph.spaceBefore = 12.5;
    paragraph.spaceAfter = 6;
    expect(paragraph.spaceBefore).toBe(12.5);
    expect(paragraph.spaceAfter).toBe(6);
    const xml = await slideXml(pptx);
    expect(xml).toContain('<a:spcBef><a:spcPts val="1250"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft>');

    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect(back.spaceBefore).toBe(12.5);
    expect(back.spaceAfter).toBe(6);

    back.spaceBefore = undefined;
    expect(back.spaceBefore).toBeUndefined();
  });

  it('keeps lnSpc before the spacing nodes it writes', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Existing', { name: 'Body' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace(
      '<a:pPr/>',
      '<a:pPr marL="457200" rtl="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst><a:defRPr sz="2400"/></a:pPr>',
    );
    zip.file('ppt/slides/slide1.xml', xml);
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    const paragraph = opened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];

    expect(paragraph.spaceBefore).toBeUndefined();
    paragraph.spaceAfter = 3;
    paragraph.leftMargin = 914400;
    const written = await slideXml(opened);
    expect(written).toMatch(
      /<a:pPr marL="914400" rtl="1"><a:lnSpc><a:spcPct val="90000"\/><\/a:lnSpc><a:spcAft><a:spcPts val="300"\/><\/a:spcAft><a:tabLst><a:tab pos="914400" algn="l"\/><\/a:tabLst><a:defRPr sz="2400"\/><\/a:pPr>/,
    );
  });

  it('reports percentage spacing as inherited and leaves it untouched', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Percent', { name: 'Body' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace('<a:pPr/>', '<a:pPr><a:spcBef><a:spcPct val="50000"/></a:spcBef></a:pPr>');
    zip.file('ppt/slides/slide1.xml', xml);
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    const paragraph = opened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    expect(paragraph.spaceBefore).toBeUndefined();
    paragraph.spaceAfter = 6;
    const written = await slideXml(opened);
    expect(written).toContain('<a:spcBef><a:spcPct val="50000"/></a:spcBef>');
    expect(written).toContain('<a:spcAft><a:spcPts val="600"/></a:spcAft>');
  });

  it('rejects non-finite point values', () => {
    const paragraph = paragraphOf('Spaced');
    for (const bad of ['12', {}, NaN, Infinity]) {
      expect(() => { paragraph.spaceBefore = bad as number; }).toThrow(/Invalid paragraph spaceBefore/);
      expect(() => { paragraph.spaceAfter = bad as number; }).toThrow(/Invalid paragraph spaceAfter/);
    }
    expect(paragraph.spaceBefore).toBeUndefined();
  });
});

describe('paragraph bullets', () => {
  it('writes buNone, buChar, and buAutoNum and round-trips them', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('a\nb\nc\nd', { name: 'List' });
    const [none, char, auto, plain] = pptx.slides.get(0).shapes.findByName('List')!.textFrame.paragraphs;

    expect(plain.bullet).toBeUndefined();

    none.bullet = 'none';
    char.bullet = { type: 'char', char: '◦' };
    auto.bullet = { type: 'autoNumber', autoNumber: 'alphaLcParen', number: 3 };
    expect(none.bullet).toEqual({ type: 'none' });
    expect(char.bullet).toEqual({ type: 'char', char: '◦' });
    expect(auto.bullet).toEqual({ type: 'autoNumber', autoNumber: 'alphaLcParen', number: 3 });

    const xml = await slideXml(pptx);
    expect(xml).toContain('<a:buNone/>');
    expect(xml).toContain('<a:buChar char="◦"/>');
    expect(xml).toContain('<a:buAutoNum type="alphaLcParen" num="3"/>');

    const reopened = await Presentation.open(await pptx.toBuffer());
    const back = reopened.slides.get(0).shapes.findByName('List')!.textFrame.paragraphs;
    expect(back[0].bullet).toEqual({ type: 'none' });
    expect(back[1].bullet).toEqual({ type: 'char', char: '◦' });
    expect(back[2].bullet).toEqual({ type: 'autoNumber', autoNumber: 'alphaLcParen', number: 3 });
    expect(back[3].bullet).toBeUndefined();
  });

  it('replaces the previous marker and keeps buFont and buClr siblings', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('• bullet', { name: 'List' });
    const paragraph = slide.shapes.findByName('List')!.textFrame.paragraphs[0];
    expect(paragraph.bullet).toEqual({ type: 'char', char: '•' });

    paragraph.bullet = 'autoNumber';
    expect(paragraph.bullet).toEqual({ type: 'autoNumber', autoNumber: 'arabicPeriod' });
    let xml = await slideXml(pptx);
    expect(xml).toContain('<a:buFont');
    expect(xml).toContain('<a:buAutoNum type="arabicPeriod"/>');
    expect(xml).not.toContain('a:buChar');

    paragraph.bullet = 'none';
    xml = await slideXml(pptx);
    expect(xml).toMatch(/<a:buFont typeface="Calibri"\/><a:buNone\/>/);

    paragraph.bullet = undefined;
    xml = await slideXml(pptx);
    expect(xml).not.toContain('a:buNone');
    expect(xml).toContain('<a:buFont typeface="Calibri"/>');
    expect(paragraph.bullet).toBeUndefined();
  });

  it('accepts bare kinds and defaults, and rejects malformed descriptors', () => {
    const paragraph = paragraphOf('x', 'List');
    paragraph.bullet = 'char';
    expect(paragraph.bullet).toEqual({ type: 'char', char: '•' });
    paragraph.bullet = { type: 'autoNumber' };
    expect(paragraph.bullet).toEqual({ type: 'autoNumber', autoNumber: 'arabicPeriod' });
    for (const bad of ['disc', '', 'constructor', 1, { type: 'char', char: '' }, { type: 'autoNumber', autoNumber: 'roman' }, { type: 'autoNumber', number: 0 }, {}, []]) {
      expect(() => { paragraph.bullet = bad as never; }).toThrow(TypeError);
      expect(() => { paragraph.bullet = bad as never; }).toThrow(/Invalid paragraph bullet|Invalid bullet/);
    }
    expect(paragraph.bullet).toEqual({ type: 'autoNumber', autoNumber: 'arabicPeriod' });
    // Like `alignment`, `null` is accepted as "inherited" and clears the marker.
    paragraph.bullet = null as never;
    expect(paragraph.bullet).toBeUndefined();
  });

  it('keeps a:buAutoNum before a:buChar position required by the schema', () => {
    const pPr: XmlNode = { 'a:buChar': { '@_char': '•' }, 'a:buFont': { '@_typeface': 'Arial' }, 'a:defRPr': '' };
    writeBullet(pPr, 'autoNumber');
    expect(Object.keys(pPr)).toEqual(['a:buFont', 'a:buAutoNum', 'a:defRPr']);
    expect(readBullet(pPr)).toEqual({ type: 'autoNumber', autoNumber: 'arabicPeriod' });
    writeBullet(pPr, undefined);
    expect(Object.keys(pPr)).toEqual(['a:buFont', 'a:defRPr']);
    expect(readBullet(pPr)).toBeUndefined();
  });

  it('reads unknown autoNumber types as the default without crashing', () => {
    const pPr: XmlNode = { 'a:buAutoNum': { '@_type': 'constructor' } };
    expect(readBullet(pPr)).toEqual({ type: 'autoNumber', autoNumber: 'arabicPeriod' });
    const bullet: ParagraphBullet = { type: 'char', char: '‣' };
    expect(bullet.type).toBe('char');
  });
});

describe('paragraph format dirty tracking', () => {
  it('does not mark the paragraph dirty when the value does not change', () => {
    let dirty = 0;
    const paragraph = new Paragraph({ 'a:pPr': { '@_marL': '457200' }, 'a:r': { 'a:t': 'x' } }, () => { dirty += 1; });

    expect(paragraph.level).toBeUndefined();
    expect(paragraph.leftMargin).toBe(457200);
    expect(paragraph.indentation).toBeUndefined();
    expect(paragraph.spaceBefore).toBeUndefined();
    expect(paragraph.bullet).toBeUndefined();
    expect(dirty).toBe(0);

    paragraph.leftMargin = 457200;
    paragraph.level = undefined;
    paragraph.spaceAfter = undefined;
    paragraph.bullet = undefined;
    expect(dirty).toBe(0);

    paragraph.level = 1;
    paragraph.indentation = -228600;
    paragraph.spaceBefore = 6;
    paragraph.bullet = 'none';
    expect(dirty).toBe(4);

    paragraph.level = 1;
    paragraph.bullet = 'none';
    expect(dirty).toBe(4);
  });

  it('keeps untouched slides byte-identical when only formatting is read', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Read only\n• bullet', { name: 'Body' });
    const before = await slideXml(pptx);
    const opened = await Presentation.open(await pptx.toBuffer());
    for (const paragraph of opened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs) {
      void [paragraph.level, paragraph.leftMargin, paragraph.indentation, paragraph.spaceBefore, paragraph.spaceAfter, paragraph.bullet];
    }
    expect(await slideXml(opened)).toBe(before);
  });

  it('inserts a newly created a:pPr as the first child of a:p', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Fresh', { name: 'Body' });
    const paragraph = pptx.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0];
    paragraph.leftMargin = 914400;
    paragraph.spaceBefore = 2;
    paragraph.bullet = { type: 'char', char: '-' };
    const xml = await slideXml(pptx);
    expect(xml).toMatch(
      /<a:p><a:pPr marL="914400"><a:spcBef><a:spcPts val="200"\/><\/a:spcBef><a:buChar char="-"\/><\/a:pPr><a:r>/,
    );
  });
});
