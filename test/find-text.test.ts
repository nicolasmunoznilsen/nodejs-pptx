import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Presentation, findTextInFrame, findTextInShape } from '../src/index.js';

async function withSplitRunText(text: string): Promise<Buffer> {
  const pptx = Presentation.create();
  pptx.slides.add().addTextBox(text, { name: 'Report' });
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  xml = xml.replace('<a:t>placeholder</a:t>', '<a:t>Informe </a:t></a:r><a:r><a:rPr b="1"/><a:t>2026 anual</a:t>');
  zip.file('ppt/slides/slide1.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('finding text across all slides', () => {
  it('reports the slide, shape, paragraph, and run of every match in presentation order', async () => {
    const created = Presentation.create();
    created.slides.add().addTextBox('Hola mundo\nAdiós mundo', { name: 'Greeting' });
    created.slides.add().addTextBox('Sin coincidencias', { name: 'Body' });
    const third = created.slides.add();
    third.addTextBox('mundo al final', { name: 'Tail' });
    third.addTextBox('mundo otra vez', { name: 'Again' });
    const pptx = await Presentation.open(await created.toBuffer());

    const matches = pptx.findText('mundo');

    expect(matches.map(m => m.matchedText)).toEqual(['mundo', 'mundo', 'mundo', 'mundo']);
    expect(matches.map(m => [m.slideIndex, m.slideNumber, m.shapeName, m.paragraphIndex, m.runIndex])).toEqual([
      [0, 1, 'Greeting', 0, 0],
      [0, 1, 'Greeting', 1, 0],
      [2, 3, 'Tail', 0, 0],
      [2, 3, 'Again', 0, 0],
    ]);
    expect(matches.map(m => m.slidePath)).toEqual([
      'ppt/slides/slide1.xml',
      'ppt/slides/slide1.xml',
      'ppt/slides/slide3.xml',
      'ppt/slides/slide3.xml',
    ]);
    expect(matches[0]).toMatchObject({ shapeType: 'textBox', runCount: 1, start: 5, end: 10, paragraphText: 'Hola mundo', runText: 'Hola mundo' });
    expect(matches[0].shape.id).toBe(matches[0].shapeId);
    expect(matches[0].shapeText).toBe('Hola mundo\nAdiós mundo');
    expect(matches[2].shapeIndex).toBe(0);
    expect(matches[3].shapeIndex).toBe(1);
  });

  it('finds text split across runs of the same paragraph', async () => {
    const pptx = await Presentation.open(await withSplitRunText('placeholder'));
    const [match] = pptx.findText('Informe 2026');
    expect(match).toMatchObject({
      slideIndex: 0,
      slideNumber: 1,
      shapeName: 'Report',
      paragraphIndex: 0,
      runIndex: 0,
      runCount: 2,
      start: 0,
      end: 12,
      matchedText: 'Informe 2026',
      paragraphText: 'Informe 2026 anual',
    });
    expect(pptx.findText('2026')[0]).toMatchObject({ runIndex: 1, runCount: 1, start: 8, end: 12 });
  });

  it('reports each occurrence with its own offsets, including repeated text in one run', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('ana vio a ana', { name: 'Mirror' });
    const matches = pptx.findText('ana');
    expect(matches.map(m => [m.start, m.end, m.matchedText])).toEqual([
      [0, 3, 'ana'],
      [10, 13, 'ana'],
    ]);
    expect(matches.every(m => m.runIndex === 0 && m.paragraphIndex === 0)).toBe(true);
  });

  it('supports regular expressions and case-insensitive string search', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Informe FY2026 y FY2027', { name: 'Body' });
    pptx.slides.add().addTextBox('informe financiero', { name: 'Title 1' });

    expect(pptx.findText(/FY\d{4}/).map(m => m.matchedText)).toEqual(['FY2026', 'FY2027']);
    expect(pptx.findText('informe').map(m => m.slideIndex)).toEqual([1]);
    expect(pptx.findText('informe', { caseSensitive: false }).map(m => m.slideIndex)).toEqual([0, 1]);
    expect(pptx.findText(/informe/g).map(m => m.slideIndex)).toEqual([1]);
    expect(pptx.findText('no-existe')).toEqual([]);

    const regex = /FY\d{4}/g;
    expect(pptx.findText(regex)).toHaveLength(2);
    expect(pptx.findText(regex)).toHaveLength(2);
    expect(regex.lastIndex).toBe(0);
  });

  it('can stop at the first match of each paragraph', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('uno uno\nuno dos', { name: 'List' });
    slide.addTextBox('uno tres', { name: 'Other' });

    expect(pptx.findText('uno', { firstOnly: true }).map(m => [m.shapeName, m.paragraphIndex, m.start])).toEqual([
      ['List', 0, 0],
      ['List', 1, 0],
      ['Other', 0, 0],
    ]);
  });

  it('exposes a slide-scoped find() and standalone helpers', async () => {
    const created = Presentation.create();
    created.slides.add().addTextBox('alpha', { name: 'First' });
    created.slides.add().addTextBox('alpha beta', { name: 'Second' });
    const pptx = await Presentation.open(await created.toBuffer());

    expect(pptx.slides.get(1).find('alpha').map(m => [m.slideIndex, m.slideNumber, m.shapeName])).toEqual([[1, 2, 'Second']]);
    expect(pptx.slides.get(1).find('alpha')[0]).toMatchObject({ slidePath: 'ppt/slides/slide2.xml' });

    const shape = pptx.slides.get(0).shapes.findByName('First')!;
    expect(findTextInShape(shape, 'alpha')).toMatchObject([{ text: 'alpha', shapeName: 'First', runIndex: 0, frame: expect.any(Object) }]);
    const bare = findTextInFrame(shape.textFrame, 'lpha');
    expect(bare).toMatchObject([{ matchedText: 'lpha', start: 1, end: 5, slideIndex: 0, slidePath: '', shapeName: '' }]);
    expect(findTextInFrame(shape.textFrame, 'alpha beta')).toEqual([]);
  });

  it('skips shapes without a text frame and returns live shapes', async () => {
    const created = Presentation.create();
    created.slides.add().addTextBox('texto', { name: 'Text' });
    const zip = await JSZip.loadAsync(await created.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace('</p:spTree>', '<p:pic><p:nvPicPr><p:cNvPr id="42" name="Picture 42"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></p:spTree>');
    zip.file('ppt/slides/slide1.xml', xml);
    const pptx = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));

    const matches = pptx.findText('texto');
    expect(matches).toHaveLength(1);
    const match = matches[0]!;
    expect(match.shapeName).toBe('Text');
    expect(match.shapeIndex).toBe(pptx.slides.get(0).shapes.all.findIndex(shape => shape.name === 'Text'));
    expect(match.shape.hasTextFrame).toBe(true);
    expect(match.frame.paragraphs[0].text).toBe('texto');
    const picture = pptx.slides.get(0).shapes.findByName('Picture 42')!;
    expect(picture.hasTextFrame).toBe(false);
    expect(findTextInShape(picture, 'texto')).toEqual([]);
  });

  it('rejects an empty search string', () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('anything');
    expect(() => pptx.findText('')).toThrow('Cannot search for an empty string.');
  });

  it('searches slides in presentation order rather than zip file order', async () => {
    const created = Presentation.create();
    created.slides.add().addTextBox('primera', { name: 'A' });
    created.slides.add().addTextBox('postrera', { name: 'B' });
    created.slides.add().addTextBox('tercera', { name: 'C' });
    const pptx = await Presentation.open(await created.toBuffer());
    pptx.slides.move(2, 0);

    expect(pptx.findText('era').map(m => [m.slideIndex, m.shapeName])).toEqual([
      [0, 'C'],
      [1, 'A'],
      [2, 'B'],
    ]);
  });
});
