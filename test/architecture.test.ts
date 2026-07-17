import JSZip from 'jszip';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Presentation, resolveRelationshipTarget, resolveRootRelationshipTarget } from '../src/index.js';

describe('OOXML path resolution', () => {
  it('resolves root relationship targets from the package root', () => {
    expect(resolveRootRelationshipTarget('ppt/presentation.xml')).toBe('ppt/presentation.xml');
    expect(resolveRootRelationshipTarget('/ppt/presentation.xml')).toBe('ppt/presentation.xml');
  });

  it('resolves absolute, relative, dot, and parent relationship targets without OS paths', () => {
    expect(resolveRelationshipTarget('ppt/presentation.xml', 'slides/slide1.xml')).toBe('ppt/slides/slide1.xml');
    expect(resolveRelationshipTarget('ppt/slides/slide1.xml', '../slideLayouts/slideLayout1.xml')).toBe('ppt/slideLayouts/slideLayout1.xml');
    expect(resolveRelationshipTarget('ppt/slideLayouts/slideLayout1.xml', '../slideMasters/slideMaster1.xml')).toBe('ppt/slideMasters/slideMaster1.xml');
    expect(resolveRelationshipTarget('ppt//slides/slide1.xml', './media/../notes/notesSlide1.xml')).toBe('ppt/slides/notes/notesSlide1.xml');
    expect(resolveRelationshipTarget('ppt/slides/slide1.xml', '/ppt/theme/theme1.xml')).toBe('ppt/theme/theme1.xml');
  });
});

describe('modular presentation behavior', () => {
  it('keeps existing public API compatible', async () => {
    const pptx = Presentation.create({ title: 'Compatibility' });
    const slide = pptx.slides.add();
    slide.addTextBox('Hello world', { name: 'Greeting' });
    expect(pptx.slides.length).toBe(1);
    expect(pptx.slides.get(0).text).toBe('Hello world');
    expect(pptx.slides.get(0).getTextBoxes()[0]).toMatchObject({ name: 'Greeting', text: 'Hello world' });
    expect(pptx.slides.get(0).replaceText('world', 'PPTX')).toBe(1);
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.get(0).text).toBe('Hello PPTX');
  });

  it('allocates a free slide part without overwriting non-consecutive existing slides', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('one');
    pptx.slides.add().addTextBox('three');
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const slide2 = await zip.file('ppt/slides/slide2.xml')!.async('string');
    const rels2 = await zip.file('ppt/slides/_rels/slide2.xml.rels')!.async('string');
    zip.file('ppt/slides/slide3.xml', slide2.replace('three', 'original-three'));
    zip.file('ppt/slides/_rels/slide3.xml.rels', rels2);
    zip.remove('ppt/slides/slide2.xml');
    zip.remove('ppt/slides/_rels/slide2.xml.rels');
    let presRels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
    presRels = presRels.replace('slides/slide2.xml', 'slides/slide3.xml');
    zip.file('ppt/_rels/presentation.xml.rels', presRels);
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    const opened = await Presentation.open(buffer);
    opened.slides.add().addTextBox('new-two');
    const out = await JSZip.loadAsync(await opened.toBuffer());
    expect(await out.file('ppt/slides/slide3.xml')!.async('string')).toContain('original-three');
    expect(await out.file('ppt/slides/slide2.xml')!.async('string')).toContain('new-two');
  });

  it('allocates shape ids from every cNvPr in the shape tree', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('first');
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace('</p:spTree>', '<p:pic><p:nvPicPr><p:cNvPr id="42" name="Picture 42"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="43" name="Chart 43"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:graphicFrame></p:spTree>');
    zip.file('ppt/slides/slide1.xml', xml);
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    opened.slides.get(0).addTextBox('after mixed shapes');
    expect(opened.slides.get(0).getTextBoxes().at(-1)?.id).toBe('44');
  });

  it('exposes shapes, text frames, paragraphs, runs, and editable run formatting', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Title\nSecond', { name: 'Title 1', fontName: 'Calibri', fontSize: 18 });
    const shape = slide.shapes.findByName('Title 1')!;
    expect(shape.hasTextFrame).toBe(true);
    expect(shape.textFrame.paragraphs).toHaveLength(2);
    const run = shape.textFrame.paragraphs[0].runs[0];
    expect(run.text).toBe('Title');
    run.text = 'Nuevo título';
    run.font.bold = true;
    run.font.italic = true;
    run.font.color = '#FF0000';
    const reopened = await Presentation.open(await pptx.toBuffer());
    const reopenedRun = reopened.slides.get(0).shapes.findByName('Title 1')!.textFrame.paragraphs[0].runs[0];
    expect(reopenedRun.text).toBe('Nuevo título');
    expect(reopenedRun.font.bold).toBe(true);
    expect(reopenedRun.font.italic).toBe(true);
    expect(reopenedRun.font.color).toBe('FF0000');
  });

  it('replaces text across multiple runs while preserving the first affected run', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('placeholder', { name: 'Report' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace('<a:t>placeholder</a:t>', '<a:t>Informe </a:t></a:r><a:r><a:rPr b="1"/><a:t>2026</a:t>');
    zip.file('ppt/slides/slide1.xml', xml);
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    expect(opened.slides.get(0).replaceText('Informe 2026', 'Informe 2027')).toBe(1);
    const reopened = await Presentation.open(await opened.toBuffer());
    expect(reopened.slides.get(0).text).toBe('Informe 2027');
    expect(reopened.slides.get(0).shapes.findByName('Report')!.textFrame.paragraphs[0].runs).toHaveLength(2);
  });

  it('preserves unmodified XML entries when saving without changes', async () => {
    const pptx = Presentation.create({ title: 'Round trip' });
    pptx.slides.add().addTextBox('No changes');
    const beforeBuffer = await pptx.toBuffer();
    const before = await JSZip.loadAsync(beforeBuffer);
    const opened = await Presentation.open(beforeBuffer);
    const after = await JSZip.loadAsync(await opened.toBuffer());
    for (const part of ['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', 'ppt/slides/slide1.xml', 'ppt/slides/_rels/slide1.xml.rels']) {
      expect(await after.file(part)!.async('string')).toBe(await before.file(part)!.async('string'));
    }
  });
});

// Keep node:fs imports exercised for environments that assert save() behavior through disk IO.
describe('save on disk', () => {
  it('saves a generated file', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-'));
    const out = join(dir, 'deck.pptx');
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('disk');
    await pptx.save(out);
    expect((await readFile(out)).length).toBeGreaterThan(0);
    await writeFile(join(dir, 'marker.txt'), 'ok');
  });
});
