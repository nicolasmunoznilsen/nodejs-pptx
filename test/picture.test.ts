import JSZip from 'jszip';
import { deflateSync, inflateSync } from 'node:zlib';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Presentation, contentTypeForExtension, detectImage, naturalImageSize, resolveRelationshipTarget } from '../src/index.js';
import type { PictureOptions } from '../src/index.js';
import { IMAGE_REL } from '../src/ooxml/constants.js';

function png(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  const raw = Buffer.alloc(height * (1 + width * 3));
  for (let row = 0; row < height; row += 1) raw[row * (1 + width * 3)] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

function chunk(type: string, data: Buffer): Buffer {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([length, body, crc]);
}

function crc32(data: Buffer): number {
  let crc = -1;
  for (const byte of data) {
    crc ^= byte;
    for (let i = 0; i < 8; i += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return ~crc >>> 0;
}

function jpeg(width: number, height: number): Buffer {
  const sof = Buffer.alloc(17);
  sof[0] = 8; sof.writeUInt16BE(height, 1); sof.writeUInt16BE(width, 3); sof[5] = 3;
  for (let component = 0; component < 3; component += 1) {
    sof[6 + component * 3] = component + 1;
    sof[7 + component * 3] = component === 0 ? 0x22 : 0x11;
    sof[8 + component * 3] = component === 0 ? 0 : 1;
  }
  const app0 = Buffer.concat([Buffer.from('JFIF\0', 'ascii'), Buffer.alloc(9)]);
  return Buffer.concat([
    Buffer.from([0xff, 0xd8, 0xff]),
    Buffer.from([0xe0, 0x00, 0x10]), app0,
    Buffer.from([0xff, 0xc0, 0x00, sof.length + 2]), sof,
    Buffer.from([0xff, 0xd9]),
  ]);
}

/**
 * Standard (non-placeable) WMF `METAHEADER`: zero type/delay words, a plausible
 * header size in words, planes and object counts in range, and a published
 * `nversion` word at offset 18.
 */
function wmfMetaheader(): Buffer {
  const header = Buffer.alloc(22);
  header.writeUInt16LE(9, 4);
  header.writeUInt16LE(1, 6);
  header.writeUInt16LE(0, 8);
  header.writeUInt16LE(0x0300, 18);
  return header;
}

describe('picture insertion', () => {
  it('detects PNG and JPEG formats and pixel sizes from magic bytes', () => {
    expect(detectImage(png(64, 32))).toMatchObject({ format: 'png', extension: 'png', contentType: 'image/png', width: 64, height: 32 });
    expect(detectImage(jpeg(320, 200))).toMatchObject({ format: 'jpeg', extension: 'jpg', contentType: 'image/jpeg', width: 320, height: 200 });
  });

  it('rejects unsupported formats naming the format', () => {
    expect(() => detectImage(Buffer.from('GIF89a\x01\x00'))).toThrow(/GIF/);
    expect(() => detectImage(Buffer.from('not an image at all'))).toThrow(/Unsupported image format/);
    expect(() => detectImage(Buffer.alloc(0))).toThrow(/Unsupported image format/);
  });

  it('writes a media part, relationship, content type and p:pic node', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const info = slide.addPicture(png(96, 48), { x: 1, y: 2, width: 3, height: 1.5, name: 'Logo' });
    expect(info).toMatchObject({ name: 'Logo', fileName: 'ppt/media/image1.png', contentType: 'image/png', format: 'png' });
    expect(info.relationshipId).toMatch(/^rId\d+$/);

    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const media = zip.file('ppt/media/image1.png');
    expect(media).not.toBeNull();
    expect(await media.async('nodebuffer')).toEqual(png(96, 48));
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"');
    expect(rels).toContain('Target="../media/image1.png"');
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).toContain('<Default Extension="png" ContentType="image/png"/>');
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).toContain('<p:pic>');
    expect(slideXml).toMatch(/<p:cNvPr id="2" name="Logo"\/>/);
    expect(slideXml).toContain('<a:blip r:embed="' + info.relationshipId + '"/>');
    expect(slideXml).toContain('<a:off x="914400" y="1828800"/>');
    expect(slideXml).toContain('<a:ext cx="2743200" cy="1371600"/>');
    expect(slideXml).toContain('<a:picLocks noChangeAspect="1"/>');
    expect(slideXml).toContain('<a:prstGeom prst="rect">');
  });

  it('defaults the display size to the natural image size at 96 DPI', async () => {
    const slide = Presentation.create().slides.add();
    const info = slide.addPicture(png(96, 48));
    expect(info.width).toBeCloseTo(1, 6);
    expect(info.height).toBeCloseTo(0.5, 6);
  });

  it('reopens pictures as discoverable picture shapes', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('caption', { name: 'Caption' });
    slide.addPicture(jpeg(300, 150), { name: 'Photo' });
    const reopened = await Presentation.open(await pptx.toBuffer());
    const shapes = reopened.slides.get(0).shapes;
    const picture = shapes.findByName('Photo');
    expect(picture?.type).toBe('picture');
    expect(picture?.hasImage).toBe(true);
    expect(picture?.imageFileName).toBe('ppt/media/image1.jpg');
    expect(picture?.imageContentType).toBe('image/jpeg');
    expect(shapes.all.map(s => s.type)).toEqual(['textBox', 'picture']);
    expect(reopened.slides.get(0).getPictures()[0]).toMatchObject({ name: 'Photo', format: 'jpeg' });
    expect(await (await JSZip.loadAsync(await reopened.toBuffer())).file('ppt/media/image1.jpg')?.async('nodebuffer')).toEqual(jpeg(300, 150));
  });

  it('allocates free media parts without overwriting existing images', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addPicture(png(10, 10), { name: 'first' });
    slide.addPicture(png(20, 20), { name: 'second' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(await zip.file('ppt/media/image1.png')?.async('nodebuffer')).toEqual(png(10, 10));
    expect(await zip.file('ppt/media/image2.png')?.async('nodebuffer')).toEqual(png(20, 20));
    zip.file('ppt/media/image1.png', Buffer.from('different bytes'));
    const reopened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    reopened.slides.get(0).addPicture(png(30, 30), { name: 'third' });
    const out = await JSZip.loadAsync(await reopened.toBuffer());
    expect(await out.file('ppt/media/image1.png')?.async('string')).toBe('different bytes');
    expect(await out.file('ppt/media/image3.png')?.async('nodebuffer')).toEqual(png(30, 30));
    const types = await out.file('[Content_Types].xml')!.async('string');
    expect(types.match(/Extension="png"/g)).toHaveLength(1);
  });

  it('accepts a filesystem path', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-picture-'));
    const file = join(dir, 'photo.jpg');
    await writeFile(file, jpeg(120, 60));
    const pptx = Presentation.create();
    const info = pptx.slides.add().addPicture(file, { name: 'FromDisk' });
    expect(info).toMatchObject({ fileName: 'ppt/media/image1.jpg', contentType: 'image/jpeg', format: 'jpeg' });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(await zip.file('ppt/media/image1.jpg')?.async('nodebuffer')).toEqual(jpeg(120, 60));
  });

  it('adds jpg and jpeg content-type defaults once and keeps round-trips clean', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addPicture(jpeg(10, 10));
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).toContain('<Default Extension="jpg" ContentType="image/jpeg"/>');
    expect(types).toContain('<Default Extension="jpeg" ContentType="image/jpeg"/>');
    expect(types.match(/Extension="rels"/g)).toHaveLength(1);
    expect(types.match(/Extension="xml"/g)).toHaveLength(1);
    const reopened = await Presentation.open(await pptx.toBuffer());
    reopened.slides.get(0).addTextBox('after image');
    expect(reopened.slides.get(0).text).toBe('after image');
    const finalTypes = await (await JSZip.loadAsync(await reopened.toBuffer())).file('[Content_Types].xml')!.async('string');
    expect(finalTypes.match(/Extension="jpg"/g)).toHaveLength(1);
  });

  it('reports shape geometry in inches before and after reopening', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('caption', { x: 1, y: 2, width: 3, height: 1.5, name: 'Caption' });
    slide.addPicture(png(96, 48), { name: 'Logo' });
    expect(slide.getTextBoxes()[0]).toMatchObject({ x: 1, y: 2, width: 3, height: 1.5 });
    expect(slide.getPictures()[0]).toMatchObject({ x: 0.5, y: 0.5, width: 1, height: 0.5 });

    const reopened = await Presentation.open(await pptx.toBuffer());
    const shapes = reopened.slides.get(0).shapes;
    expect(shapes.findByName('Caption')).toMatchObject({ x: 1, y: 2, width: 3, height: 1.5 });
    expect(shapes.findByName('Logo')).toMatchObject({ x: 0.5, y: 0.5, width: 1, height: 0.5 });
    expect(reopened.slides.get(0).getPictures()[0]).toMatchObject({ x: 0.5, y: 0.5, width: 1, height: 0.5 });
  });

  it('names empty input in the error', () => {
    expect(() => Presentation.create().slides.add().addPicture(Buffer.alloc(0))).toThrow(/Unsupported image format: empty or missing image data/);
  });
  it('falls back to the 640x480 default size when no header dimensions exist', () => {
    const headerless = Buffer.from([0xff, 0xd8, 0xff, 0xd9]);
    const image = detectImage(headerless);
    expect(image).toMatchObject({ format: 'jpeg', extension: 'jpg', defaultExtensions: ['jpg', 'jpeg'] });
    expect(image.width).toBeUndefined();
    expect(naturalImageSize(image)).toMatchObject({ width: 640 / 96, height: 480 / 96 });
    const info = Presentation.create().slides.add().addPicture(headerless);
    expect(info.width).toBeCloseTo(640 / 96, 6);
    expect(info.height).toBeCloseTo(480 / 96, 6);
    expect(contentTypeForExtension('.JPG')).toBe('image/jpeg');
  });

  it('reports an actionable error for unsupported files', () => {
    const slide = Presentation.create().slides.add();
    expect(() => slide.addPicture(Buffer.from('GIF87a\x00\x01'))).toThrow(/Unsupported image format: GIF/);
  });

  it('names EMF and WMF records with their real byte order', () => {
    const emf = Buffer.alloc(88);
    emf.writeUInt32LE(1, 0);
    emf.writeUInt32LE(emf.length, 4);
    const wmf = Buffer.from([0x9a, 0xc6, 0xcd, 0xd7, 0, 0, 0, 0]);
    expect(() => detectImage(emf)).toThrow(/Unsupported image format: EMF/);
    expect(() => detectImage(wmf)).toThrow(/Unsupported image format: WMF/);
    expect(() => detectImage(Buffer.from([0xd7, 0xcd, 0xc6, 0x9a, 0, 0, 0, 0]))).toThrow(/Unsupported image format: WMF/);
  });

  it('names TIFF, unknown and standard WMF METAHEADER payloads', async () => {
    const cases: Array<[Buffer, string]> = [
      [Buffer.from([0x49, 0x49, 0x2a, 0x00, 0x04, 0x00]), 'TIFF'],
      [Buffer.from([0x4d, 0x4d, 0x00, 0x2a, 0x00, 0x04]), 'TIFF'],
      [Buffer.from([1, 2, 3, 4, 5]), 'unknown'],
      [wmfMetaheader(), 'WMF'],
    ];
    for (const [payload, format] of cases) {
      expect(() => detectImage(payload)).toThrow(new RegExp(`^Unsupported image format: ${format}`));
      const pptx = Presentation.create();
      const slide = pptx.slides.add();
      expect(() => slide.addPicture(payload)).toThrow(new RegExp(`^Unsupported image format: ${format}`));
      expect(mediaNames(await JSZip.loadAsync(await pptx.toBuffer()))).toEqual([]);
    }
  });

  it('rejects truncated PNG and JPEG payloads', () => {
    const full = png(32, 16);
    expect(() => detectImage(full.subarray(0, 12))).toThrow(/Unsupported image format: truncated PNG/);
    expect(() => detectImage(full.subarray(0, full.length - 12))).toThrow(/Unsupported image format: truncated PNG/);
    const truncatedJpeg = jpeg(32, 16).subarray(0, 12);
    expect(() => detectImage(truncatedJpeg)).toThrow(/Unsupported image format: truncated JPEG/);
    expect(() => Presentation.create().slides.add().addPicture(jpeg(32, 16).subarray(0, -2))).toThrow(/Unsupported image format: truncated JPEG/);
  });

  it('accepts plain Uint8Array views', async () => {
    const image = png(64, 32);
    const padded = Buffer.concat([Buffer.alloc(4), image]);
    const view = new Uint8Array(padded.buffer, padded.byteOffset + 4, image.length);
    const pptx = Presentation.create();
    const info = pptx.slides.add().addPicture(view, { name: 'View' });
    expect(info).toMatchObject({ name: 'View', fileName: 'ppt/media/image1.png', contentType: 'image/png', format: 'png' });
    expect(info.width).toBeCloseTo(64 / 96, 6);
    expect(info.height).toBeCloseTo(32 / 96, 6);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(await zip.file('ppt/media/image1.png')?.async('nodebuffer')).toEqual(image);
  });

  it('keeps media part names unique across formats and registers each default once', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const first = slide.addPicture(png(10, 10));
    const second = slide.addPicture(jpeg(10, 10));
    const third = slide.addPicture(png(20, 20));
    expect([first.fileName, second.fileName, third.fileName]).toEqual(['ppt/media/image1.png', 'ppt/media/image2.jpg', 'ppt/media/image3.png']);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const types = await zip.file('[Content_Types].xml')!.async('string');
    for (const ext of ['png', 'jpg', 'jpeg']) expect(types.match(new RegExp(`Extension="${ext}"`, 'g'))).toHaveLength(1);
    expect(await zip.file('ppt/media/image1.png')?.async('nodebuffer')).toEqual(png(10, 10));
    expect(await zip.file('ppt/media/image2.jpg')?.async('nodebuffer')).toEqual(jpeg(10, 10));
    expect(await zip.file('ppt/media/image3.png')?.async('nodebuffer')).toEqual(png(20, 20));
  });

  it('leaves the slide and package unchanged when the payload is unsupported', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('keep me', { name: 'Caption' });
    const before = slide.getPictures().length;
    const payloads = [
      Buffer.from('GIF89a\x01\x00'),
      Buffer.from('BM\x36\x00\x00\x00'),
      Buffer.from('II*\x00\x04\x00'),
      Buffer.concat([Buffer.from('RIFF'), Buffer.alloc(4), Buffer.from('WEBPVP8 ')]),
      Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>'),
      Buffer.from([0x01, 0x00, 0x00, 0x00, 0x58, 0x00, 0x00, 0x00]),
      Buffer.from([0xd7, 0xcd, 0xc6, 0x9a, 0, 0, 0, 0]),
      Buffer.from([0x89, 0x50, 0x4e, 0x47]),
      png(8, 8).subarray(0, 30),
    ];
    for (const payload of payloads) {
      await expect(async () => slide.addPicture(payload)).rejects.toThrow(/Unsupported image format/);
    }
    expect(slide.getPictures()).toHaveLength(before);
    expect(slide.shapes.all.map(shape => shape.type)).toEqual(['textBox']);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(Object.keys(zip.files).filter(name => name.startsWith('ppt/media/'))).toEqual([]);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).not.toContain('image/png');
    expect(types).not.toContain('image/jpeg');
    expect(await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).not.toContain('/image');
  });

  it('reads every picture back after save and reopen', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-picture-reopen-'));
    const file = join(dir, 'photo.jpeg');
    await writeFile(file, jpeg(120, 60));
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const first = slide.addPicture(png(96, 48), { x: 1, y: 1, width: 2, height: 1, name: 'Logo' });
    const second = slide.addPicture(file, { name: 'Photo' });
    const deck = join(dir, 'deck.pptx');
    await pptx.save(deck);
    const pictures = (await Presentation.open(deck)).slides.get(0).getPictures();
    expect(pictures.map(picture => picture.index)).toEqual([0, 1]);
    expect(pictures.map(picture => picture.name)).toEqual(['Logo', 'Photo']);
    expect(pictures.map(picture => picture.fileName)).toEqual(['ppt/media/image1.png', 'ppt/media/image2.jpg']);
    expect(pictures.map(picture => picture.contentType)).toEqual(['image/png', 'image/jpeg']);
    expect(pictures.map(picture => picture.format)).toEqual(['png', 'jpeg']);
    expect(pictures[0]).toMatchObject({ x: 1, y: 1, width: 2, height: 1, relationshipId: first.relationshipId });
    expect(pictures[1].width).toBeCloseTo(1.25, 6);
    expect(pictures[1].relationshipId).toBe(second.relationshipId);
    const rels = await (await JSZip.loadAsync(await readFile(deck))).file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
    const imageRel = (id: string, target: string): string => `Id="${id}" Type="${IMAGE_REL}" Target="${target}"`;
    expect(rels).toContain(imageRel(first.relationshipId, '../media/image1.png'));
    expect(rels).toContain(imageRel(second.relationshipId, '../media/image2.jpg'));
  });
});

/** Independent acceptance verification for issue #21 (PNG/JPEG insertion). */
const EMU_PER_INCH_ACCEPTANCE = 914400;
const emuOf = (inches: number): number => Math.round(inches * EMU_PER_INCH_ACCEPTANCE);

function acceptancePng(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const raw = Buffer.alloc(height * (1 + width * 3), 0x7f);
  for (let row = 0; row < height; row += 1) raw[row * (1 + width * 3)] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

function acceptanceJpeg(width: number, height: number): Buffer {
  const marker = (code: number, body: Buffer): Buffer => {
    const head = Buffer.alloc(2);
    head.writeUInt16BE(body.length + 2, 0);
    return Buffer.concat([Buffer.from([0xff, code]), head, body]);
  };
  const jfif = Buffer.concat([Buffer.from('JFIF\0', 'ascii'), Buffer.from([1, 1, 1, 0, 0, 96, 0, 96, 0, 0])]);
  const sof = Buffer.concat([
    Buffer.from([8]),
    (() => { const b = Buffer.alloc(2); b.writeUInt16BE(height, 0); return b; })(),
    (() => { const b = Buffer.alloc(2); b.writeUInt16BE(width, 0); return b; })(),
    Buffer.from([3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1]),
  ]);
  const dqt = Buffer.concat([Buffer.from([0]), Buffer.alloc(64, 8)]);
  const dht = Buffer.concat([Buffer.from([0]), Buffer.alloc(16, 1), Buffer.alloc(11, 2)]);
  const sos = Buffer.concat([Buffer.from([3, 1, 0, 2, 0, 3, 0, 0, 0x3f, 0]), Buffer.from([0x2a, 0x55, 0xaa])]);
  return Buffer.concat([
    Buffer.from([0xff, 0xd8]),
    marker(0xe0, jfif),
    marker(0xdb, dqt),
    marker(0xc0, sof),
    marker(0xc4, dht),
    marker(0xda, sos),
    Buffer.from([0xff, 0xd9]),
  ]);
}

function defaultEntries(contentTypes: string): Array<{ extension: string; contentType: string }> {
  return [...contentTypes.matchAll(/<Default\b[^>]*?>/g)].map((match) => ({
    extension: /Extension="([^"]*)"/.exec(match[0])?.[1] ?? '',
    contentType: /ContentType="([^"]*)"/.exec(match[0])?.[1] ?? '',
  }));
}

function packageEntries(zip: JSZip): string[] {
  const names: string[] = [];
  zip.forEach((path, entry) => { if (!entry.dir) names.push(path); });
  return names.sort();
}

async function pictureRelationships(zip: JSZip, slidePart: string): Promise<Array<{ id: string; type: string; target: string }>> {
  const rels = await zip.file(relsPathOf(slidePart))!.async('string');
  return [...rels.matchAll(/<Relationship\b[^>]*?>/g)].map((match) => ({
    id: /Id="([^"]*)"/.exec(match[0])?.[1] ?? '',
    type: /Type="([^"]*)"/.exec(match[0])?.[1] ?? '',
    target: /Target="([^"]*)"/.exec(match[0])?.[1] ?? '',
  }));
}

async function embeddedIds(zip: JSZip, slidePart: string): Promise<string[]> {
  const slideXml = await zip.file(slidePart)!.async('string');
  return [...slideXml.matchAll(/<a:blip\b[^>]*?r:embed="([^"]*)"/g)].map(match => match[1]);
}

async function transformOf(zip: JSZip, slidePart: string, pictureIndex: number): Promise<{ x: number; y: number; cx: number; cy: number }> {
  const slideXml = await zip.file(slidePart)!.async('string');
  const pic = [...slideXml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)][pictureIndex]?.[0] ?? '';
  const off = /<a:off\b[^>]*?x="(-?\d+)"[^>]*?y="(-?\d+)"/.exec(pic);
  const ext = /<a:ext\b[^>]*?cx="(\d+)"[^>]*?cy="(\d+)"/.exec(pic);
  expect(off, `a:off in picture ${pictureIndex}`).not.toBeNull();
  expect(ext, `a:ext in picture ${pictureIndex}`).not.toBeNull();
  return { x: Number(off![1]), y: Number(off![2]), cx: Number(ext![1]), cy: Number(ext![2]) };
}

function relsPathOf(partName: string): string {
  const i = partName.lastIndexOf('/');
  return `${partName.slice(0, i)}/_rels/${partName.slice(i + 1)}.rels`;
}

describe('issue #21 acceptance: PNG and JPEG insertion', () => {
  it('acceptance fixtures are structurally valid images', () => {
    const pngBytes = acceptancePng(24, 12);
    const idatStart = pngBytes.indexOf('IDAT', 8, 'ascii') + 4;
    const idatLength = pngBytes.readUInt32BE(idatStart - 4);
    const raw = inflateSync(pngBytes.subarray(idatStart, idatStart + idatLength));
    expect(raw.length).toBe(12 * (1 + 24 * 3));
    const jpegBytes = acceptanceJpeg(24, 12);
    expect([...jpegBytes.subarray(0, 2)]).toEqual([0xff, 0xd8]);
    expect(jpegBytes.indexOf(Buffer.from([0xff, 0xda]))).toBeGreaterThan(0);
    expect([...jpegBytes.subarray(jpegBytes.length - 2)]).toEqual([0xff, 0xd9]);
    const sof = jpegBytes.indexOf(Buffer.from([0xff, 0xc0]));
    expect(sof).toBeGreaterThan(0);
    expect(jpegBytes.readUInt16BE(sof + 5)).toBe(12);
    expect(jpegBytes.readUInt16BE(sof + 7)).toBe(24);
  });

  it('PNG insert: stores byte-identical media and exactly one png Default', async () => {
    const pngBytes = acceptancePng(120, 80);
    const pptx = Presentation.create();
    pptx.slides.add().addPicture(pngBytes);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(packageEntries(zip).filter(name => name.startsWith('ppt/media/'))).toEqual(['ppt/media/image1.png']);
    const stored = await zip.file('ppt/media/image1.png')!.async('nodebuffer');
    expect(stored).toHaveLength(pngBytes.length);
    expect(Buffer.compare(stored, pngBytes)).toBe(0);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    const pngDefaults = defaultEntries(types).filter(entry => entry.extension.toLowerCase() === 'png');
    expect(pngDefaults).toHaveLength(1);
    expect(pngDefaults[0].contentType).toBe('image/png');
  });

  it('JPEG insert: jpg/jpeg Defaults are unique and r:embed resolves to the stored media path', async () => {
    const jpegBytes = acceptanceJpeg(200, 100);
    const pptx = Presentation.create();
    pptx.slides.add().addPicture(jpegBytes);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const slidePart = 'ppt/slides/slide1.xml';
    const types = await zip.file('[Content_Types].xml')!.async('string');
    const defaults = defaultEntries(types);
    expect(defaults.map(entry => entry.extension)).toEqual([...new Set(defaults.map(entry => entry.extension))]);
    for (const extension of ['jpg', 'jpeg']) {
      const matches = defaults.filter(entry => entry.extension.toLowerCase() === extension);
      expect(matches).toHaveLength(1);
      expect(matches[0].contentType).toBe('image/jpeg');
    }
    const embeds = await embeddedIds(zip, slidePart);
    expect(embeds).toHaveLength(1);
    const imageRels = await pictureRelationships(zip, slidePart);
    const resolved = imageRels.filter(rel => rel.id === embeds[0]);
    expect(resolved).toHaveLength(1);
    expect(resolved[0].type).toBe(IMAGE_REL);
    const mediaPath = resolveRelationshipTarget(slidePart, resolved[0].target);
    expect(mediaPath).toBe('ppt/media/image1.jpg');
    expect(packageEntries(zip)).toContain(mediaPath);
    expect(Buffer.compare(await zip.file(mediaPath)!.async('nodebuffer'), jpegBytes)).toBe(0);
  });

  it('two images of the same format get distinct media parts and distinct relationships', async () => {
    const first = acceptancePng(10, 10);
    const validSecond = acceptancePng(30, 30);
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const a = slide.addPicture(first);
    const b = slide.addPicture(validSecond);
    expect(a.fileName).not.toBe(b.fileName);
    expect(a.relationshipId).not.toBe(b.relationshipId);
    expect([a.fileName, b.fileName]).toEqual(['ppt/media/image1.png', 'ppt/media/image2.png']);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(packageEntries(zip).filter(name => name.startsWith('ppt/media/'))).toEqual(['ppt/media/image1.png', 'ppt/media/image2.png']);
    expect(Buffer.compare(await zip.file('ppt/media/image1.png')!.async('nodebuffer'), first)).toBe(0);
    expect(Buffer.compare(await zip.file('ppt/media/image2.png')!.async('nodebuffer'), validSecond)).toBe(0);
    const embeds = await embeddedIds(zip, 'ppt/slides/slide1.xml');
    expect(embeds).toHaveLength(2);
    expect(new Set(embeds).size).toBe(2);
    expect(embeds).toEqual([a.relationshipId, b.relationshipId]);
  });

  it('explicit x/y/width/height become a:xfrm EMU values of inches()*914400', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const info = slide.addPicture(acceptancePng(120, 80), { x: 0.5, y: 1.25, width: 3, height: 2 });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(await transformOf(zip, 'ppt/slides/slide1.xml', 0)).toEqual({
      x: emuOf(0.5), y: emuOf(1.25), cx: emuOf(3), cy: emuOf(2),
    });
    expect(info).toMatchObject({ x: 0.5, y: 1.25, width: 3, height: 2 });
    expect(emuOf(0.5)).toBe(457200);
    expect(emuOf(3)).toBe(2743200);
  });

  it('omitted sizes fall back to PNG IHDR and JPEG SOFn pixels at 96 DPI', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const pngInfo = slide.addPicture(acceptancePng(120, 80));
    const jpegInfo = slide.addPicture(acceptanceJpeg(200, 100));
    expect(pngInfo.width).toBeCloseTo(120 / 96, 9);
    expect(pngInfo.height).toBeCloseTo(80 / 96, 9);
    expect(jpegInfo.width).toBeCloseTo(200 / 96, 9);
    expect(jpegInfo.height).toBeCloseTo(100 / 96, 9);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(await transformOf(zip, 'ppt/slides/slide1.xml', 0)).toEqual({
      x: emuOf(0.5), y: emuOf(0.5), cx: emuOf(120 / 96), cy: emuOf(80 / 96),
    });
    expect(await transformOf(zip, 'ppt/slides/slide1.xml', 1)).toEqual({
      x: emuOf(0.5), y: emuOf(0.5), cx: emuOf(200 / 96), cy: emuOf(100 / 96),
    });
  });

  it('unsupported bytes reject with /Unsupported image format/ leaving the package untouched', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addPicture(acceptancePng(16, 16));
    const beforeBuffer = await pptx.toBuffer();
    const beforeZip = await JSZip.loadAsync(beforeBuffer);
    const beforeNames = packageEntries(beforeZip);
    const beforeSlide = await beforeZip.file('ppt/slides/slide1.xml')!.async('string');
    const beforeRels = await beforeZip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
    const beforeTypes = await beforeZip.file('[Content_Types].xml')!.async('string');
    expect(slide.getPictures()).toHaveLength(1);
    const badPayloads: Array<[string, Buffer]> = [
      ['GIF8 header', Buffer.from('GIF89a\x01\x00\x01\x00\x00\x00')],
      ['GIF87a header', Buffer.from('GIF87a\x01\x00\x01\x00')],
      ['truncated 2-byte PNG', acceptancePng(16, 16).subarray(0, 2)],
      ['truncated 2-byte JPEG', Buffer.from([0xff, 0xd8])],
      ['single byte', Buffer.from([0x42])],
      ['empty', Buffer.alloc(0)],
    ];
    for (const [, payload] of badPayloads) {
      expect(() => slide.addPicture(payload)).toThrow(/Unsupported image format/);
    }
    expect(slide.getPictures()).toHaveLength(1);
    const afterBuffer = await pptx.toBuffer();
    const afterZip = await JSZip.loadAsync(afterBuffer);
    expect(packageEntries(afterZip)).toEqual(beforeNames);
    expect(await afterZip.file('ppt/slides/slide1.xml')!.async('string')).toBe(beforeSlide);
    expect(await afterZip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).toBe(beforeRels);
    expect(await afterZip.file('[Content_Types].xml')!.async('string')).toBe(beforeTypes);
    expect((await embeddedIds(afterZip, 'ppt/slides/slide1.xml'))).toHaveLength(1);
  });

  it('getPictures() after save and reopen reports format/fileName/contentType/relationshipId', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-issue21-'));
    const deck = join(dir, 'deck.pptx');
    const pngBytes = acceptancePng(120, 80);
    const jpegBytes = acceptanceJpeg(200, 100);
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const written = [slide.addPicture(pngBytes, { name: 'Logo' }), slide.addPicture(jpegBytes, { name: 'Photo' })];
    await pptx.save(deck);
    const reopened = await Presentation.open(await readFile(deck));
    const pictures = reopened.slides.get(0).getPictures();
    expect(pictures.map(picture => picture.name)).toEqual(['Logo', 'Photo']);
    expect(pictures.map(picture => picture.format)).toEqual(['png', 'jpeg']);
    expect(pictures.map(picture => picture.fileName)).toEqual(['ppt/media/image1.png', 'ppt/media/image2.jpg']);
    expect(pictures.map(picture => picture.contentType)).toEqual(['image/png', 'image/jpeg']);
    expect(pictures.map(picture => picture.relationshipId)).toEqual(written.map(info => info.relationshipId));
    expect(new Set(pictures.map(picture => picture.relationshipId)).size).toBe(2);
    const zip = await JSZip.loadAsync(await readFile(deck));
    const rels = await pictureRelationships(zip, 'ppt/slides/slide1.xml');
    for (const [index, picture] of pictures.entries()) {
      const rel = rels.filter(candidate => candidate.id === picture.relationshipId);
      expect(rel).toHaveLength(1);
      expect(rel[0].type).toBe(IMAGE_REL);
      expect(resolveRelationshipTarget('ppt/slides/slide1.xml', rel[0].target)).toBe(picture.fileName);
      expect(Buffer.compare(await zip.file(picture.fileName)!.async('nodebuffer'), index === 0 ? pngBytes : jpegBytes)).toBe(0);
    }
    expect(pictures[0]).toMatchObject({ x: 0.5, y: 0.5, width: 120 / 96, height: 80 / 96 });
  });
});

/** Round 2 acceptance: reopen-safe media allocation, anchored error text, path inputs. */
function mediaNames(zip: JSZip): string[] {
  return packageEntries(zip).filter(name => name.startsWith('ppt/media/'));
}

async function imageRelsOf(zip: JSZip, slidePart: string): Promise<Array<{ id: string; type: string; target: string }>> {
  return (await pictureRelationships(zip, slidePart)).filter(rel => rel.type === IMAGE_REL);
}

function imageDefaultsOf(contentTypes: string): string[] {
  return defaultEntries(contentTypes)
    .filter(entry => ['png', 'jpg', 'jpeg'].includes(entry.extension.toLowerCase()))
    .map(entry => `${entry.extension.toLowerCase()}=${entry.contentType}`)
    .sort();
}

describe('issue #21 acceptance: reopen, error text and path inputs', () => {
  it('keeps media names unique and defaults single across save, reopen and re-add', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-issue21-reopen-'));
    const deck = join(dir, 'deck.pptx');
    const pngA = acceptancePng(20, 10);
    const jpegA = acceptanceJpeg(30, 15);
    const pngB = acceptancePng(40, 20);
    const jpegB = acceptanceJpeg(50, 25);
    const pptx = Presentation.create();
    const first = pptx.slides.add();
    first.addPicture(pngA, { name: 'one' });
    first.addPicture(jpegA, { name: 'two' });
    await pptx.save(deck);

    const reopened = await Presentation.open(deck);
    const second = reopened.slides.add();
    const third = reopened.slides.get(0).addPicture(pngB, { name: 'three' });
    const fourth = second.addPicture(jpegB, { name: 'four' });
    expect([third.fileName, fourth.fileName]).toEqual(['ppt/media/image3.png', 'ppt/media/image4.jpg']);
    await reopened.save(deck);

    const zip = await JSZip.loadAsync(await readFile(deck));
    expect(mediaNames(zip)).toEqual([
      'ppt/media/image1.png', 'ppt/media/image2.jpg', 'ppt/media/image3.png', 'ppt/media/image4.jpg',
    ]);
    expect(Buffer.compare(await zip.file('ppt/media/image1.png')!.async('nodebuffer'), pngA)).toBe(0);
    expect(Buffer.compare(await zip.file('ppt/media/image2.jpg')!.async('nodebuffer'), jpegA)).toBe(0);
    expect(Buffer.compare(await zip.file('ppt/media/image3.png')!.async('nodebuffer'), pngB)).toBe(0);
    expect(Buffer.compare(await zip.file('ppt/media/image4.jpg')!.async('nodebuffer'), jpegB)).toBe(0);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(imageDefaultsOf(types)).toEqual(['jpeg=image/jpeg', 'jpg=image/jpeg', 'png=image/png']);
    for (const [slidePart, expected] of [['ppt/slides/slide1.xml', 3], ['ppt/slides/slide2.xml', 1]] as Array<[string, number]>) {
      const rels = await imageRelsOf(zip, slidePart);
      expect(rels).toHaveLength(expected);
      expect(new Set(rels.map(rel => rel.id)).size).toBe(expected);
      expect(new Set(rels.map(rel => resolveRelationshipTarget(slidePart, rel.target))).size).toBe(expected);
    }
    expect(await reopened.slides.get(0).getPictures()).toHaveLength(3);
    expect(reopened.slides.get(0).getPictures().map(picture => picture.fileName))
      .toEqual(['ppt/media/image1.png', 'ppt/media/image2.jpg', 'ppt/media/image3.png']);
  });

  it('keeps png/jpg/jpeg Defaults exactly once when reopening a deck that already declares them', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addPicture(acceptancePng(12, 6));
    pptx.slides.add().addPicture(acceptanceJpeg(24, 12));
    const reopened = await Presentation.open(await pptx.toBuffer());
    const reopenedTypes = await (await JSZip.loadAsync(await reopened.toBuffer())).file('[Content_Types].xml')!.async('string');
    expect(imageDefaultsOf(reopenedTypes)).toEqual(['jpeg=image/jpeg', 'jpg=image/jpeg', 'png=image/png']);
    reopened.slides.get(0).addPicture(acceptancePng(30, 15));
    reopened.slides.get(0).addPicture(acceptanceJpeg(60, 30));
    const types = await (await JSZip.loadAsync(await reopened.toBuffer())).file('[Content_Types].xml')!.async('string');
    expect(imageDefaultsOf(types)).toEqual(['jpeg=image/jpeg', 'jpg=image/jpeg', 'png=image/png']);
    for (const extension of ['png', 'jpg', 'jpeg']) {
      expect(types.match(new RegExp(`Extension="${extension}"`, 'g'))).toHaveLength(1);
    }
  });

  it('errors start with Unsupported image format and name the detected format', async () => {
    const cases: Array<[string, Buffer]> = [
      ['GIF', Buffer.from('GIF89a\x01\x00\x01\x00\x00\x00')],
      ['BMP', Buffer.concat([Buffer.from('BM'), Buffer.alloc(60)])],
      ['WEBP', Buffer.concat([Buffer.from('RIFF'), Buffer.alloc(4), Buffer.from('WEBPVP8 '), Buffer.alloc(8)])],
      ['SVG', Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>')],
      ['empty or missing image data', Buffer.alloc(0)],
      ['truncated PNG', acceptancePng(16, 16).subarray(0, acceptancePng(16, 16).length - 12)],
      ['truncated JPEG', acceptanceJpeg(16, 16).subarray(0, acceptanceJpeg(16, 16).indexOf(0xff, 4))],
    ];
    for (const [format, payload] of cases) {
      const slide = Presentation.create().slides.add();
      expect(() => slide.addPicture(payload)).toThrow(
        new RegExp(`^Unsupported image format: ${format.replace(/[ /]/g, '\\$&')}`),
      );
    }
  });

  it('leaves the package byte-for-byte unchanged after every rejected payload', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addPicture(acceptancePng(16, 16), { name: 'keep' });
    const before = await JSZip.loadAsync(await pptx.toBuffer());
    const beforeNames = packageEntries(before);
    const beforeMedia = mediaNames(before);
    const beforeRels = await before.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
    const beforeTypes = await before.file('[Content_Types].xml')!.async('string');
    const beforeSlide = await before.file('ppt/slides/slide1.xml')!.async('string');
    const beforeCount = slide.getPictures().length;
    const payloads = [
      Buffer.from('GIF89a\x01\x00\x01\x00\x00\x00'),
      Buffer.concat([Buffer.from('BM'), Buffer.alloc(60)]),
      Buffer.concat([Buffer.from('RIFF'), Buffer.alloc(4), Buffer.from('WEBPVP8 '), Buffer.alloc(8)]),
      Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>'),
      Buffer.alloc(0),
      acceptancePng(16, 16).subarray(0, acceptancePng(16, 16).length - 12),
      acceptanceJpeg(16, 16).subarray(0, acceptanceJpeg(16, 16).indexOf(0xff, 4)),
    ];
    for (const payload of payloads) {
      expect(() => slide.addPicture(payload)).toThrow(/^Unsupported image format: /);
      const mid = await JSZip.loadAsync(await pptx.toBuffer());
      expect(mediaNames(mid)).toEqual(beforeMedia);
      expect(await mid.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).toBe(beforeRels);
      expect(await mid.file('[Content_Types].xml')!.async('string')).toBe(beforeTypes);
    }
    expect(slide.getPictures()).toHaveLength(beforeCount);
    const after = await JSZip.loadAsync(await pptx.toBuffer());
    expect(packageEntries(after)).toEqual(beforeNames);
    expect(mediaNames(after)).toEqual(beforeMedia);
    expect(await after.file('ppt/slides/slide1.xml')!.async('string')).toBe(beforeSlide);
    expect(await after.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).toBe(beforeRels);
    expect(await after.file('[Content_Types].xml')!.async('string')).toBe(beforeTypes);
    expect(await embeddedIds(after, 'ppt/slides/slide1.xml')).toHaveLength(beforeCount);
  });

  it('round-trips Buffer, Uint8Array and path bytes byte-identically through save and reopen', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-issue21-path-'));
    const pngPath = join(dir, 'logo.png');
    const jpegPath = join(dir, 'photo.jpg');
    const pngBytes = acceptancePng(120, 80);
    const jpegBytes = acceptanceJpeg(200, 100);
    await writeFile(pngPath, pngBytes);
    await writeFile(jpegPath, jpegBytes);
    const padded = Buffer.concat([Buffer.alloc(7, 0xab), pngBytes, Buffer.alloc(5, 0xcd)]);
    const view = new Uint8Array(padded.buffer, padded.byteOffset + 7, pngBytes.length);
    expect(view.byteOffset).toBeGreaterThan(0);

    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const bufferPng = slide.addPicture(pngBytes, { name: 'bufferPng' });
    const bufferJpeg = slide.addPicture(jpegBytes, { name: 'bufferJpeg' });
    const viewPng = slide.addPicture(view, { name: 'viewPng' });
    const pathPng = slide.addPicture(pngPath, { name: 'pathPng' });
    const pathJpeg = slide.addPicture(jpegPath, { name: 'pathJpeg' });
    expect([bufferPng.fileName, bufferJpeg.fileName, viewPng.fileName, pathPng.fileName, pathJpeg.fileName])
      .toEqual(['ppt/media/image1.png', 'ppt/media/image2.jpg', 'ppt/media/image3.png', 'ppt/media/image4.png', 'ppt/media/image5.jpg']);
    const deck = join(dir, 'deck.pptx');
    await pptx.save(deck);

    const reopened = await Presentation.open(deck);
    const savedZip = await JSZip.loadAsync(await readFile(deck));
    const expected: Array<[string, Buffer]> = [
      ['ppt/media/image1.png', pngBytes],
      ['ppt/media/image2.jpg', jpegBytes],
      ['ppt/media/image3.png', pngBytes],
      ['ppt/media/image4.png', pngBytes],
      ['ppt/media/image5.jpg', jpegBytes],
    ];
    for (const [part, bytes] of expected) {
      expect(Buffer.compare(await savedZip.file(part)!.async('nodebuffer'), bytes)).toBe(0);
    }
    expect(mediaNames(savedZip)).toEqual(expected.map(([part]) => part));
    const resaved = await JSZip.loadAsync(await reopened.toBuffer());
    for (const [part, bytes] of expected) {
      expect(Buffer.compare(await resaved.file(part)!.async('nodebuffer'), bytes)).toBe(0);
    }
    expect(reopened.slides.get(0).getPictures().map(picture => picture.format))
      .toEqual(['png', 'jpeg', 'png', 'png', 'jpeg']);

    await expect(async () => slide.addPicture(join(dir, 'missing.png'))).rejects.toThrow(/ENOENT/);
    const afterMissing = await JSZip.loadAsync(await pptx.toBuffer());
    expect(mediaNames(afterMissing)).toEqual(expected.map(([part]) => part));
    expect(await afterMissing.file('ppt/slides/_rels/slide1.xml.rels')!.async('string'))
      .toBe(await savedZip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string'));
    expect(slide.getPictures()).toHaveLength(expected.length);
  });
});

describe('issue #21 acceptance: spPr geometry and picture count', () => {
  it('stores p:pic spPr xfrm EMU values that match getPictures() and the p:pic count', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('caption', { name: 'Caption' });
    const explicit = slide.addPicture(acceptancePng(120, 80), { x: 1.25, y: 2.5, width: 3, height: 1.5, name: 'Explicit' });
    const natural = slide.addPicture(acceptanceJpeg(192, 96), { name: 'Natural' });
    expect([explicit.width, explicit.height]).toEqual([3, 1.5]);
    expect(natural.width).toBeCloseTo(192 / 96, 9);
    expect(natural.height).toBeCloseTo(96 / 96, 9);

    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    const pics = [...slideXml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)].map(match => match[0]);
    expect(pics).toHaveLength(2);
    const expected = [
      { x: 1.25, y: 2.5, cx: 3, cy: 1.5 },
      { x: 0.5, y: 0.5, cx: 192 / 96, cy: 96 / 96 },
    ];
    for (const [index, pic] of pics.entries()) {
      const spPr = /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(pic)?.[0] ?? '';
      expect(spPr, `p:spPr of picture ${index}`).toContain('<a:xfrm>');
      expect(/<a:xfrm><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><\/a:xfrm>/.test(spPr), `xfrm in p:spPr of picture ${index}`).toBe(true);
      const off = /<a:off\b[^>]*?x="(-?\d+)"[^>]*?y="(-?\d+)"/.exec(spPr)!;
      const ext = /<a:ext\b[^>]*?cx="(\d+)"[^>]*?cy="(\d+)"/.exec(spPr)!;
      expect([Number(off[1]), Number(off[2]), Number(ext[1]), Number(ext[2])]).toEqual([
        emuOf(expected[index].x), emuOf(expected[index].y), emuOf(expected[index].cx), emuOf(expected[index].cy),
      ]);
    }
    expect(slide.getPictures()).toHaveLength(pics.length);
    expect(slide.shapes.all.filter(shape => shape.type === 'picture')).toHaveLength(pics.length);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const pictures = reopened.slides.get(0).getPictures();
    expect(pictures).toHaveLength(pics.length);
    expect(pictures.map(picture => [picture.x, picture.y, picture.width, picture.height])).toEqual([
      [1.25, 2.5, 3, 1.5],
      [0.5, 0.5, 192 / 96, 96 / 96],
    ]);
    expect(pictures.map(picture => picture.index)).toEqual([0, 1]);
    expect(pictures.map(picture => picture.name)).toEqual(['Explicit', 'Natural']);
    expect(pictures.map(picture => picture.format)).toEqual(['png', 'jpeg']);
    expect(pictures.map(picture => picture.fileName)).toEqual(['ppt/media/image1.png', 'ppt/media/image2.jpg']);
    expect(pictures.map(picture => picture.contentType)).toEqual(['image/png', 'image/jpeg']);
    expect(new Set(pictures.map(picture => picture.relationshipId)).size).toBe(2);
  });
});

/** Independent acceptance verification for issue #37 (picture geometry validation). */
const GEOMETRY_ERROR = /picture (x|y|width|height)/;

/** JSZip 3.x has no fileNames(); Object.keys(zip.files) is the equivalent entry listing. */
function zipFileNames(zip: JSZip): string[] {
  const names: string[] = [];
  zip.forEach((path, entry) => { if (!entry.dir) names.push(path); });
  return names.sort();
}

function geometryZipState(pptx: Presentation): Promise<{ media: string[]; slideXml: string; rels: string }> {
  return (async () => {
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    return {
      media: zipFileNames(zip).filter(name => name.includes('ppt/media')).sort(),
      slideXml: await zip.file('ppt/slides/slide1.xml')!.async('string'),
      rels: await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string'),
    };
  })();
}

describe('issue #37 acceptance: picture position and size validation', () => {
  it('addPicture throws a TypeError naming the offending option for invalid geometry', () => {
    const slide = Presentation.create().slides.add();
    const badPositions: Array<[string, unknown]> = [
      ['x', NaN], ['y', NaN],
      ['x', Infinity], ['y', Infinity], ['x', -Infinity], ['y', -Infinity],
      ['x', '1'], ['y', '2'],
    ];
    for (const [name, value] of badPositions) {
      expect(() => slide.addPicture(png(24, 12), { [name]: value as number }), `x/y=${String(value)}`).toThrow(TypeError);
      expect(() => slide.addPicture(png(24, 12), { [name]: value as number }), `x/y=${String(value)}`).toThrow(GEOMETRY_ERROR);
    }
    const badSizes: Array<[string, unknown]> = [
      ['width', 0], ['height', 0],
      ['width', -1], ['height', -0.5], ['width', -1e9],
      ['width', NaN], ['height', NaN], ['width', Infinity], ['height', -Infinity],
      ['width', '3'], ['height', '1.5'],
    ];
    for (const [name, value] of badSizes) {
      expect(() => slide.addPicture(png(24, 12), { [name]: value as number }), `${name}=${String(value)}`).toThrow(TypeError);
      expect(() => slide.addPicture(png(24, 12), { [name]: value as number }), `${name}=${String(value)}`).toThrow(GEOMETRY_ERROR);
    }
  });

  it('leaves media parts, slide XML and rels untouched and keeps the picture count when rejecting', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('keep me', { name: 'Caption' });
    slide.addPicture(png(32, 16), { name: 'keep' });
    const before = await geometryZipState(pptx);
    const beforeCount = slide.getPictures().length;
    expect(before.media).toEqual(['ppt/media/image1.png']);
    expect(beforeCount).toBe(1);

    const invalidOptions: PictureOptions[] = [
      { x: NaN }, { y: Infinity }, { x: '1' as unknown as number },
      { width: 0 }, { height: -2 }, { width: NaN }, { height: '1.5' as unknown as number },
      { x: 1, y: 1, width: -3, height: 1 },
    ];
    for (const options of invalidOptions) {
      expect(() => slide.addPicture(jpeg(40, 20), options), JSON.stringify(options)).toThrow(TypeError);
      const after = await geometryZipState(pptx);
      expect(after.media, `media after ${JSON.stringify(options)}`).toEqual(before.media);
      expect(after.slideXml, `slide xml after ${JSON.stringify(options)}`).toBe(before.slideXml);
      expect(after.rels, `rels after ${JSON.stringify(options)}`).toBe(before.rels);
      expect(slide.getPictures().length, `count after ${JSON.stringify(options)}`).toBe(beforeCount);
    }
  });

  it('validates geometry before inspecting the payload so an unsupported buffer reports geometry first', () => {
    const slide = Presentation.create().slides.add();
    for (const payload of [Buffer.from('GIF89a\x01\x00\x01\x00\x00\x00'), Buffer.alloc(0), Buffer.from('not an image at all')]) {
      expect(() => slide.addPicture(payload, { x: NaN })).toThrow(TypeError);
      expect(() => slide.addPicture(payload, { x: NaN })).toThrow(/picture x/);
      expect(() => slide.addPicture(payload, { width: 0 })).toThrow(/picture width/);
      let message = '';
      try { slide.addPicture(payload, { height: -1 }); } catch (error) { message = (error as Error).message; }
      expect(message).toMatch(/picture height/);
      expect(message).not.toMatch(/Unsupported image format/);
    }
    expect(slide.getPictures()).toHaveLength(0);
  });

  it('keeps negative and omitted geometry working and stores negative offsets outside the canvas', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const omitted = slide.addPicture(png(96, 48), { name: 'omitted' });
    expect(omitted.x).toBe(0.5);
    expect(omitted.y).toBe(0.5);
    expect(omitted.width).toBeCloseTo(1, 6);
    expect(omitted.height).toBeCloseTo(0.5, 6);

    const negative = slide.addPicture(png(96, 48), { x: -1, y: -1, width: 3, height: 1.5 });
    expect([negative.x, negative.y, negative.width, negative.height]).toEqual([-1, -1, 3, 1.5]);

    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).toContain('<a:off x="-914400" y="-914400"/>');
    expect(slideXml).toContain('<a:ext cx="2743200" cy="1371600"/>');

    const reopened = await Presentation.open(await pptx.toBuffer());
    const pictures = reopened.slides.get(0).getPictures();
    expect(pictures.map(picture => [picture.x, picture.y, picture.width, picture.height])).toEqual([
      [0.5, 0.5, 1, 0.5],
      [-1, -1, 3, 1.5],
    ]);
  });

  it('still accepts a valid call right after a rejected one', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    expect(() => slide.addPicture(png(24, 12), { x: NaN, name: 'rejected' })).toThrow(/picture x/);
    expect(slide.getPictures()).toHaveLength(0);
    const info = slide.addPicture(jpeg(48, 24), { x: 2, y: 3, width: 4, height: 2, name: 'valid' });
    expect(info).toMatchObject({ name: 'valid', fileName: 'ppt/media/image1.jpg', width: 4, height: 2, x: 2, y: 3 });
    expect(slide.getPictures().map(picture => picture.name)).toEqual(['valid']);
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    expect(zipFileNames(zip).filter(name => name.includes('ppt/media')).sort()).toEqual(['ppt/media/image1.jpg']);
    expect(await zip.file('ppt/media/image1.jpg')!.async('nodebuffer')).toEqual(jpeg(48, 24));
    const pictures = (await Presentation.open(await pptx.toBuffer())).slides.get(0).getPictures();
    expect(pictures.map(picture => picture.name)).toEqual(['valid']);
  });
});
