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
