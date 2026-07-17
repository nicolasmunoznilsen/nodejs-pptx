import JSZip from 'jszip';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/index.js';

describe('Presentation', () => {
  it('creates, reopens, edits, saves, and preserves unknown OOXML files', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-'));
    const out = join(dir, 'deck.pptx');
    const edited = join(dir, 'edited.pptx');

    const created = Presentation.create({ title: 'Demo' });
    const slide = created.slides.add();
    slide.addTextBox('Hello world', { x: 1, y: 1, width: 4, height: 1, name: 'Greeting' });
    const buffer = await created.toBuffer();

    const zip = await JSZip.loadAsync(buffer);
    zip.file('ppt/custom/unknown.xml', '<custom><value>preserve me</value></custom>');
    await BunlessWrite(out, await zip.generateAsync({ type: 'nodebuffer' }));

    const opened = await Presentation.open(await readFile(out));
    expect(opened.slides.length).toBe(1);
    expect(opened.slides.get(0).text).toBe('Hello world');
    expect(opened.slides.get(0).getTextBoxes()[0]).toMatchObject({ text: 'Hello world', name: 'Greeting', index: 0 });
    opened.slides.get(0).addTextBox('Second text', { x: 2, y: 2, width: 3, height: 1 });
    expect(opened.slides.get(0).replaceText('Hello', 'Hola')).toBe(1);
    await opened.save(edited);

    const reopened = await Presentation.open(edited);
    expect(reopened.slides.get(0).text).toContain('Hola world');
    expect(reopened.slides.get(0).text).toContain('Second text');
    const finalZip = await JSZip.loadAsync(await readFile(edited));
    expect(await finalZip.file('ppt/custom/unknown.xml')?.async('string')).toBe('<custom><value>preserve me</value></custom>');
  });

  it('throws clear errors for invalid input and slide indexes', async () => {
    await expect(Presentation.open('/not/found/example.pptx')).rejects.toThrow(/does not exist/);
    await expect(Presentation.open(Buffer.from('not a zip'))).rejects.toThrow(/Invalid PPTX zip package/);
    const pptx = Presentation.create();
    expect(() => pptx.slides.get(0)).toThrow(/out of range/);
  });
});

async function BunlessWrite(path: string, data: Buffer): Promise<void> {
  const { writeFile } = await import('node:fs/promises');
  await writeFile(path, data);
}

describe('Slide operations, layouts, placeholders, and metadata', () => {
  it('removes first, middle, and last slides while updating package parts', async () => {
    for (const index of [0, 1, 2]) {
      const pptx = Presentation.create();
      pptx.slides.add().addTextBox('one');
      pptx.slides.add().addTextBox('two');
      pptx.slides.add().addTextBox('three');
      const removed = pptx.slides.remove(index);
      expect(removed.text).toBe(['one', 'two', 'three'][index]);
      const zip = await JSZip.loadAsync(await pptx.toBuffer());
      expect(zip.file(removed.path)).toBeNull();
      expect(zip.file(`ppt/slides/_rels/${removed.path.split('/').pop()}.rels`)).toBeNull();
      const reopened = await Presentation.open(await pptx.toBuffer());
      expect(reopened.slides.toArray().map(s => s.text)).toEqual(['one', 'two', 'three'].filter((_, i) => i !== index));
    }
  });

  it('moves slides by reordering p:sldIdLst only and survives reopen', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('one');
    pptx.slides.add().addTextBox('two');
    pptx.slides.add().addTextBox('three');
    const before = await JSZip.loadAsync(await pptx.toBuffer());
    const slideXmlBefore = await before.file('ppt/slides/slide1.xml')?.async('string');
    pptx.slides.move(2, 0);
    const after = await JSZip.loadAsync(await pptx.toBuffer());
    expect(await after.file('ppt/slides/slide1.xml')?.async('string')).toBe(slideXmlBefore);
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.toArray().map(s => s.text)).toEqual(['three', 'one', 'two']);
  });

  it('duplicates slides with XML, text runs, and relationships', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Hello');
    slide.addTextBox('World');
    const copy = pptx.slides.duplicate(0);
    expect(copy.path).not.toBe(slide.path);
    expect(pptx.slides.toArray().map(s => s.text)).toEqual(['Hello\nWorld', 'Hello\nWorld']);
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.length).toBe(2);
    expect(reopened.slides.get(1).layout?.name).toBe('Blank');
  });

  it('exposes layouts, placeholders, backgrounds, dimensions, and core properties', async () => {
    const pptx = Presentation.create({ title: 'Initial', author: 'Author' });
    expect(pptx.slideLayouts.length).toBeGreaterThan(0);
    const blank = pptx.slideLayouts.findByName('Blank');
    expect(blank?.type).toBe('blank');
    const slide = pptx.slides.add(blank);
    expect(slide.layout?.name).toBe('Blank');
    expect(slide.placeholders.length).toBe(0);
    slide.background.color = '#FF0000';
    pptx.width = 11;
    pptx.height = 6;
    pptx.title = 'Changed';
    pptx.subject = 'Subject';
    pptx.keywords = 'pptx,test';
    pptx.comments = 'Comments';
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.width).toBe(11);
    expect(reopened.height).toBe(6);
    expect(reopened.title).toBe('Changed');
    expect(reopened.subject).toBe('Subject');
    expect(reopened.keywords).toBe('pptx,test');
    expect(reopened.comments).toBe('Comments');
    expect(reopened.slides.get(0).background.color).toBe('FF0000');
  });
});
