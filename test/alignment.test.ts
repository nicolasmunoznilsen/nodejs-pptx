import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Presentation, type ParagraphAlignment } from '../src/index.js';
import { ALIGNMENT_TO_OOXML, OOXML_TO_ALIGNMENT, isAlignment } from '../src/text/alignment.js';

async function slideXml(pptx: Presentation): Promise<string> {
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  return (await zip.file('ppt/slides/slide1.xml')!.async('string'));
}

describe('paragraph text alignment', () => {
  const cases: Array<[ParagraphAlignment, string]> = [
    ['left', 'l'],
    ['center', 'ctr'],
    ['right', 'r'],
    ['justify', 'just'],
  ];

  it('writes the OOXML algn value for every supported alignment', async () => {
    for (const [name, ooxml] of cases) {
      const pptx = Presentation.create();
      const slide = pptx.slides.add();
      slide.addTextBox('Centrado', { name: 'Body', alignment: name });
      const xml = await slideXml(pptx);
      expect(xml).toContain(`algn="${ooxml}"`);
      const reopened = await Presentation.open(await pptx.toBuffer());
      expect(reopened.slides.get(0).shapes.findByName('Body')!.textFrame.paragraphs[0].alignment).toBe(name);
    }
  });

  it('reads, edits, clears, and round-trips alignment through the paragraph', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('First\nSecond', { name: 'Notes' });
    const shape = slide.shapes.findByName('Notes')!;
    const [first, second] = shape.textFrame.paragraphs;

    expect(first.alignment).toBeUndefined();
    expect(shape.textFrame.alignment).toBeUndefined();

    first.alignment = 'center';
    expect(first.alignment).toBe('center');
    expect(second.alignment).toBeUndefined();
    expect(shape.textFrame.alignment).toBeUndefined();

    shape.textFrame.alignment = 'justify';
    expect(first.alignment).toBe('justify');
    expect(second.alignment).toBe('justify');
    expect(shape.textFrame.alignment).toBe('justify');

    second.alignment = undefined;
    expect(second.alignment).toBeUndefined();

    const xml = await slideXml(pptx);
    expect(xml).toContain('<a:pPr algn="just"/>');
    expect((xml.match(/algn="just"/g) ?? []).length).toBe(1);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const reopenedParagraphs = reopened.slides.get(0).shapes.findByName('Notes')!.textFrame.paragraphs;
    expect(reopenedParagraphs[0].alignment).toBe('justify');
    expect(reopenedParagraphs[1].alignment).toBeUndefined();
    // Mixed alignments inside one frame report `undefined` at frame level.
    expect(reopened.slides.get(0).getTextBoxes()[0].alignment).toBeUndefined();
  });

  it('keeps a:pPr as the first child and preserves its other attributes', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('• bullet\nplain', { name: 'List' });
    const [bullet, plain] = slide.shapes.findByName('List')!.textFrame.paragraphs;

    bullet.alignment = 'right';
    plain.alignment = 'center';
    expect(bullet.alignment).toBe('right');
    expect(plain.alignment).toBe('center');
    const xml = await slideXml(pptx);
    expect(xml).toMatch(/<a:pPr marL="274320" indent="-274320" algn="r">/);
    // A newly created a:pPr must be the first child of a:p, before a:r.
    expect(xml).toMatch(/<a:p><a:pPr algn="ctr"\/><a:r>/);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const reopenedBullet = reopened.slides.get(0).shapes.findByName('List')!.textFrame.paragraphs[0];
    expect(reopenedBullet.alignment).toBe('right');
    expect(reopenedBullet.text).toBe('bullet');
  });

  it('reads OOXML alignment written by other tools', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('from PowerPoint', { name: 'External' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace('<a:pPr/>', '<a:pPr algn="ctr" rtl="0"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr>');
    zip.file('ppt/slides/slide1.xml', xml);
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    const paragraph = opened.slides.get(0).shapes.findByName('External')!.textFrame.paragraphs[0];
    expect(paragraph.alignment).toBe('center');

    paragraph.alignment = 'left';
    const edited = await slideXml(opened);
    expect(edited).toContain('<a:pPr algn="l" rtl="0">');
    expect(edited).toContain('<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>');
  });

  it('does not mark the slide dirty when the alignment does not change', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('stable', { name: 'Stable', alignment: 'center' });
    const before = await pptx.toBuffer();
    const opened = await Presentation.open(before);
    const shape = opened.slides.get(0).shapes.findByName('Stable')!;
    shape.textFrame.paragraphs[0].alignment = 'center';
    shape.textFrame.alignment = 'center';
    shape.textFrame.paragraphs[0].alignment = 'center';
    expect(await opened.toBuffer()).toEqual(before);
  });

  it('rejects unsupported alignment values', () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const shape = slide.addTextBox('nope', { name: 'Nope' });
    const paragraph = slide.shapes.findById(shape.id)!.textFrame.paragraphs[0];
    expect(() => { paragraph.alignment = 'middle' as ParagraphAlignment; }).toThrow(/Invalid paragraph alignment/);
    expect(() => slide.addTextBox('x', { alignment: 'both' as ParagraphAlignment })).toThrow(/Invalid paragraph alignment/);
  });

  it('rejects inherited property names as alignment values', () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const shape = slide.addTextBox('proto', { name: 'Proto' });
    const paragraph = slide.shapes.findById(shape.id)!.textFrame.paragraphs[0];
    const textFrame = slide.shapes.findById(shape.id)!.textFrame;
    const inherited = [
      'toString', 'constructor', 'hasOwnProperty', 'valueOf', 'isPrototypeOf', '__proto__', 'propertyIsEnumerable',
    ];

    for (const name of inherited) {
      expect(() => { paragraph.alignment = name as ParagraphAlignment; }).toThrow(/Invalid paragraph alignment/);
      expect(() => { textFrame.alignment = name as ParagraphAlignment; }).toThrow(/Invalid paragraph alignment/);
      expect(() => slide.addTextBox('x', { alignment: name as ParagraphAlignment })).toThrow(/Invalid paragraph alignment/);
      expect(paragraph.alignment).toBeUndefined();
    }

    // Nothing leaked into the XML and the document still serializes.
    expect(() => pptx.toBuffer()).not.toThrow();
  });

  it('never resolves an inherited name through the alignment lookup tables', async () => {
    for (const name of ['toString', 'constructor', 'hasOwnProperty', '__proto__', 'valueOf']) {
      expect(ALIGNMENT_TO_OOXML[name as ParagraphAlignment]).toBeUndefined();
      expect(OOXML_TO_ALIGNMENT[name]).toBeUndefined();
      expect(isAlignment(name)).toBe(false);
    }
    expect(isAlignment('center')).toBe(true);
    expect(isAlignment('__proto__')).toBe(false);

    // An unknown `algn` value coming from another tool reads as inherited, not as a function.
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('external', { name: 'External' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace('<a:pPr/>', '<a:pPr algn="dist"/>');
    zip.file('ppt/slides/slide1.xml', xml);
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    const paragraph = opened.slides.get(0).shapes.findByName('External')!.textFrame.paragraphs[0];
    expect(paragraph.alignment).toBeUndefined();
    expect(typeof paragraph.alignment).not.toBe('function');
    expect(opened.slides.get(0).getTextBoxes()[0].alignment).toBeUndefined();
  });

  it('leaves untouched slides byte-identical when reading alignment', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('read only', { name: 'ReadOnly', alignment: 'right' });
    const before = await pptx.toBuffer();
    const opened = await Presentation.open(before);
    expect(opened.slides.get(0).shapes.findByName('ReadOnly')!.textFrame.alignment).toBe('right');
    expect(await opened.toBuffer()).toEqual(before);
  });
});
