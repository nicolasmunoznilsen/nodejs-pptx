import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/index.js';

/** Adds a text-bearing shape to the shape tree of slide `slideNumber` (1-based) of a generated deck. */
async function injectShape(pptx: Presentation, slideNumber: number, shapeXml: string, part = `ppt/slides/slide${slideNumber}.xml`): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  const xml = await zip.file(part)!.async('string');
  zip.file(part, xml.replace('</p:spTree>', `${shapeXml}</p:spTree>`));
  return zip.generateAsync({ type: 'nodebuffer' });
}

function textBox(id: string, name: string, runsXml: string): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${runsXml}</a:p></p:txBody></p:sp>`;
}

function run(text: string, bold = false): string {
  return `<a:r>${bold ? '<a:rPr b="1"/>' : ''}<a:t>${text}</a:t></a:r>`;
}

describe('replace text across an entire presentation', () => {
  it('replaces on every slide in presentation order and returns the total count', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Report FY2026 summary');
    const second = pptx.slides.add();
    second.addTextBox('FY2026 revenue');
    second.addTextBox('FY2026 costs and FY2026 notes');
    pptx.slides.add();

    expect(pptx.replaceText('FY2026', 'FY2027')).toBe(4);

    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.get(0).text).toBe('Report FY2027 summary');
    expect(reopened.slides.get(1).text).toBe('FY2027 revenue\nFY2027 costs and FY2027 notes');
    expect(reopened.slides.get(2).text).toBe('');
  });

  it('applies regex replacements to every slide', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('quarter 1 FY2024');
    pptx.slides.add().addTextBox('quarter 4 FY2025');

    expect(pptx.replaceText(/FY20\d\d/g, 'FY2030')).toBe(2);
    expect(await (await Presentation.open(await pptx.toBuffer())).slides.toArray().map(s => s.text)).toEqual(['quarter 1 FY2030', 'quarter 4 FY2030']);
  });

  it('replaces text split across runs on every slide', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('seed');
    pptx.slides.add().addTextBox('seed');
    const first = await injectShape(pptx, 1, textBox('10', 'Split 1', `${run('Informe ')}${run('2026', true)}`));
    const zip = await JSZip.loadAsync(first);
    const secondXml = await zip.file('ppt/slides/slide2.xml')!.async('string');
    zip.file('ppt/slides/slide2.xml', secondXml.replace('</p:spTree>', `${textBox('11', 'Split 2', `${run('Informe ')}${run('2026')}`)}</p:spTree>`));
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));

    expect(opened.replaceText('Informe 2026', 'Informe 2027')).toBe(2);
    const reopened = await Presentation.open(await opened.toBuffer());
    expect(reopened.slides.toArray().map(slide => slide.text)).toEqual(['seed\nInforme 2027', 'seed\nInforme 2027']);
    // The first affected run keeps its formatting.
    const runs = reopened.slides.get(0).shapes.findByName('Split 1')!.textFrame.paragraphs[0].runs;
    expect(runs.map(r => r.text)).toEqual(['Informe 2027', '']);
    expect(runs[0].font.bold).toBe(false);
    expect(runs[1].font.bold).toBe(true);
  });

  it('reaches placeholders and shapes nested inside groups on every slide', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('seed');
    pptx.slides.add().addTextBox('seed');
    const withPlaceholder = await injectShape(pptx, 1, `<p:sp><p:nvSpPr><p:cNvPr id="20" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${run('title FY2026')}</a:p></p:txBody></p:sp>`);
    const grouped = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="30" name="Group 30"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:spTree>${textBox('31', 'Inside group', run('grouped FY2026'))}</p:spTree></p:grpSp>`;
    const zip = await JSZip.loadAsync(withPlaceholder);
    zip.file('ppt/slides/slide2.xml', (await zip.file('ppt/slides/slide2.xml')!.async('string')).replace('</p:spTree>', `${grouped}</p:spTree>`));
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));

    expect(opened.replaceText('FY2026', 'FY2027')).toBe(2);
    const reopened = await Presentation.open(await opened.toBuffer());
    expect(reopened.slides.get(0).placeholders.findByName('Title 1')!.textFrame.text).toBe('title FY2027');
    expect(reopened.slides.get(1).shapes.allDescendants.find(shape => shape.name === 'Inside group')!.textFrame.text).toBe('grouped FY2027');
  });

  it('leaves layout parts untouched by default and rewrites them with includeLayouts', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('body FY2026');
    const withLayoutText = await injectShape(pptx, 1, textBox('40', 'Layout text', run('template FY2026')), 'ppt/slideLayouts/slideLayout1.xml');
    const opened = await Presentation.open(withLayoutText);

    expect(opened.replaceText('FY2026', 'FY2027')).toBe(1);
    const untouched = await JSZip.loadAsync(await opened.toBuffer());
    expect(await untouched.file('ppt/slideLayouts/slideLayout1.xml')!.async('string')).toBe(await (await JSZip.loadAsync(withLayoutText)).file('ppt/slideLayouts/slideLayout1.xml')!.async('string'));

    const withLayouts = await Presentation.open(withLayoutText);
    expect(withLayouts.replaceText('FY2026', 'FY2027', { includeLayouts: true })).toBe(2);
    const rewritten = await Presentation.open(await withLayouts.toBuffer());
    expect(rewritten.slideLayouts.toArray()[0].shapes.allDescendants.find(shape => shape.name === 'Layout text')!.textFrame.text).toBe('template FY2027');
    expect(rewritten.slides.get(0).text).toBe('body FY2027');
  });

  it('returns 0 and preserves unmodified parts when nothing matches', async () => {
    const pptx = Presentation.create({ title: 'No match' });
    pptx.slides.add().addTextBox('Nothing to change');
    const before = await JSZip.loadAsync(await pptx.toBuffer());
    const opened = await Presentation.open(await before.generateAsync({ type: 'nodebuffer' }));

    expect(opened.replaceText('absent', 'whatever')).toBe(0);
    const after = await JSZip.loadAsync(await opened.toBuffer());
    for (const part of ['ppt/presentation.xml', 'ppt/slides/slide1.xml', 'ppt/slideLayouts/slideLayout1.xml', 'docProps/core.xml']) {
      expect(await after.file(part)!.async('string')).toBe(await before.file(part)!.async('string'));
    }
  });

  it('keeps slide.replaceText() behavior unchanged', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('Hello world');
    pptx.slides.add().addTextBox('Hello elsewhere');
    expect(pptx.slides.get(0).replaceText('Hello', 'Hola')).toBe(1);
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.toArray().map(slide => slide.text)).toEqual(['Hola world', 'Hello elsewhere']);
  });
});
