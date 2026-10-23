import JSZip from 'jszip';
import { deflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/index.js';
import { naturalImageSize, pictureXml, validatePictureGeometry } from '../src/shapes/picture.js';
import { EMU_PER_INCH } from '../src/ooxml/constants.js';

/** Minimal valid PNG with a given pixel size, matching the fixtures in test/picture.test.ts. */
function png(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; ihdr[9] = 2;
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

/** Every entry of a package buffer as raw bytes, so a later snapshot can be compared byte for byte. */
async function snapshot(buffer: Buffer): Promise<Map<string, Buffer>> {
  const zip = await JSZip.loadAsync(buffer);
  const entries = new Map<string, Buffer>();
  for (const name of Object.keys(zip.files).sort()) {
    const file = zip.files[name];
    if (file.dir) continue;
    entries.set(name, await file.async('nodebuffer'));
  }
  return entries;
}

function expectIdentical(before: Map<string, Buffer>, after: Map<string, Buffer>): void {
  expect([...after.keys()]).toEqual([...before.keys()]);
  for (const [name, bytes] of before) {
    const current = after.get(name);
    expect(current, name).toBeDefined();
    expect(current!.equals(bytes), name).toBe(true);
  }
}

/** Largest inch value whose rounded EMU is still 0, and the smallest one that is 1 EMU. */
const ZERO_EMU_LIMIT = 0.5 / EMU_PER_INCH;

describe('issue #41: a size that rounds to 0 EMU is rejected', () => {
  it('rejects positive width and height values that round to 0 EMU', () => {
    expect(Math.round(1e-9 * EMU_PER_INCH)).toBe(0);
    expect(ZERO_EMU_LIMIT).toBeCloseTo(5.4677e-7, 10);
    for (const value of [Number.MIN_VALUE, 5e-7, 1e-9, 1e-300, ZERO_EMU_LIMIT * 0.999]) {
      expect(() => validatePictureGeometry({ width: value }), String(value)).toThrow(TypeError);
      expect(() => validatePictureGeometry({ width: value }), String(value))
        .toThrow(new RegExp(`^Invalid picture width: expected a finite number greater than 0 whose EMU value is a safe integer, got ${value}\\.$`));
      expect(() => validatePictureGeometry({ height: value }), String(value))
        .toThrow(new RegExp(`^Invalid picture height: expected a finite number greater than 0 whose EMU value is a safe integer, got ${value}\\.$`));
    }
  });

  it('keeps the smallest legal size and every size that rounds to at least one EMU', () => {
    expect(() => validatePictureGeometry({ width: ZERO_EMU_LIMIT })).not.toThrow();
    expect(() => validatePictureGeometry({ height: 1 / EMU_PER_INCH })).not.toThrow();
    expect(() => validatePictureGeometry({ width: 0.001, height: 0.001 })).not.toThrow();
    expect(() => validatePictureGeometry({ width: 1e-6, height: 1e-6 })).not.toThrow();
    expect(() => validatePictureGeometry({ width: 0.5, height: 0.5 })).not.toThrow();
  });

  it('keeps rejecting zero and negative sizes with the positive-number message', () => {
    expect(() => validatePictureGeometry({ width: 0 })).toThrow(/^Invalid picture width: expected a finite number greater than 0, got 0\./);
    expect(() => validatePictureGeometry({ height: -1e-9 })).toThrow(/^Invalid picture height: expected a finite number greater than 0, got -1e-9\./);
    expect(() => validatePictureGeometry({ width: Number.NaN })).toThrow(/^Invalid picture width: expected a finite number greater than 0, got NaN\./);
  });

  it('leaves x and y semantics untouched, including values that round to 0 EMU', () => {
    expect(() => validatePictureGeometry({ x: 0, y: 0 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -0, y: 1e-9 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: Number.MIN_VALUE, y: -5e-7 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -1e-9 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -12.5, y: -0.25 })).not.toThrow();
  });
});

describe('issue #41: every geometry write path is guarded', () => {
  it('rejects a 0 EMU size coming from the pictureXml size fallback', () => {
    expect(() => pictureXml('2', 'rId2', { height: 1 }, { width: 1e-9, height: 1 })).toThrow(
      /^Invalid picture width: expected a finite number greater than 0 whose EMU value is a safe integer, got 1e-9\.$/,
    );
    expect(() => pictureXml('2', 'rId2', { width: 0.5 }, { width: 4, height: 5e-7 })).toThrow(
      /^Invalid picture height: expected a finite number greater than 0 whose EMU value is a safe integer, got 5e-7\.$/,
    );
    expect(() => pictureXml('2', 'rId2', { width: 1e-9 })).toThrow(/^Invalid picture width: /);
    expect(() => pictureXml('2', 'rId2', { x: 1e-9, y: 0 }, { width: 4, height: 3 })).not.toThrow();
    expect(() => pictureXml('2', 'rId2')).not.toThrow();
    expect(JSON.stringify(pictureXml('2', 'rId2', { x: 0, width: 2 }))).toContain('"@_cx":"1828800"');
  });

  it('never lets a natural image size round to 0 EMU', () => {
    expect(naturalImageSize({ format: 'png', extension: 'png', contentType: 'image/png', defaultExtensions: ['png'], buffer: Buffer.alloc(0), width: 1, height: 1 }))
      .toMatchObject({ width: 1 / 96, height: 1 / 96 });
    expect(Math.round((1 / 96) * EMU_PER_INCH)).toBe(9525);
    expect(naturalImageSize({ format: 'png', extension: 'png', contentType: 'image/png', defaultExtensions: ['png'], buffer: Buffer.alloc(0) }))
      .toMatchObject({ width: 640 / 96, height: 480 / 96 });
    expect(() => validatePictureGeometry(naturalImageSize({
      format: 'png', extension: 'png', contentType: 'image/png', defaultExtensions: ['png'], buffer: Buffer.alloc(0), width: 1, height: 1,
    }))).not.toThrow();
  });
});

describe('issue #41: addPicture rejects a 0 EMU size before touching the package', () => {
  it('leaves media, slide rels, [Content_Types].xml and spTree byte-identical', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('keep me', { name: 'Caption' });
    const before = await snapshot(await pptx.toBuffer());
    expect([...before.keys()].filter(name => name.startsWith('ppt/media/'))).toEqual([]);

    for (const options of [{ width: 1e-9 }, { height: 1e-9 }, { width: Number.MIN_VALUE }, { height: 5e-7 }]) {
      expect(() => slide.addPicture(png(96, 48), options), JSON.stringify(options)).toThrow(TypeError);
      expect(() => slide.addPicture(png(96, 48), options), JSON.stringify(options))
        .toThrow(/^Invalid picture (width|height): expected a finite number greater than 0 whose EMU value is a safe integer, got /);
    }

    const after = await snapshot(await pptx.toBuffer());
    expectIdentical(before, after);
    expect([...after.keys()].filter(name => name.startsWith('ppt/media/'))).toEqual([]);
    expect((after.get('[Content_Types].xml') ?? Buffer.alloc(0)).toString()).not.toMatch(/image\/(png|jpeg)/);
    expect((after.get('ppt/slides/_rels/slide1.xml.rels') ?? Buffer.alloc(0)).toString()).not.toContain('/image');
    const slideXml = (after.get('ppt/slides/slide1.xml') ?? Buffer.alloc(0)).toString();
    expect(slideXml).not.toContain('<p:pic>');
    expect(slideXml).toContain('keep me');
    expect(slide.getPictures()).toHaveLength(0);
  });

  it('rejects a 0 EMU size through pictureXml and leaves the package byte-identical', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('keep me', { name: 'Caption' });
    const before = await snapshot(await pptx.toBuffer());

    for (const value of [Number.MIN_VALUE, 5e-7, 1e-9]) {
      expect(() => pictureXml('2', 'rId2', { width: value, height: 2 }), String(value)).toThrow(TypeError);
      expect(() => pictureXml('2', 'rId2', { width: value, height: 2 }), String(value))
        .toThrow(new RegExp(`^Invalid picture width: expected a finite number greater than 0 whose EMU value is a safe integer, got ${value}\\.$`));
      expect(() => pictureXml('2', 'rId2', { width: 2, height: value }), String(value)).toThrow(TypeError);
      expect(() => pictureXml('2', 'rId2', { width: 2, height: value }), String(value))
        .toThrow(new RegExp(`^Invalid picture height: expected a finite number greater than 0 whose EMU value is a safe integer, got ${value}\\.$`));
      expect(() => slide.addPicture(png(96, 48), { width: value, height: value }), String(value)).toThrow(TypeError);
    }

    const after = await snapshot(await pptx.toBuffer());
    expectIdentical(before, after);
    expect([...after.keys()].filter(name => name.startsWith('ppt/media/'))).toEqual([]);
    expect([...after.keys()].filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))).toEqual(['ppt/slides/slide1.xml']);
    expect((after.get('ppt/slides/slide1.xml') ?? Buffer.alloc(0)).toString()).not.toContain('<p:pic>');
  });

  it('accepts negative and zero positions plus the smallest size that rounds above 0 EMU', () => {
    expect(() => validatePictureGeometry({ x: 0, y: 0 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -1 })).not.toThrow();
    expect(() => validatePictureGeometry({ y: -2.5 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -1, y: -2.5 })).not.toThrow();

    // 1.5/EMU_PER_INCH inches is ~1.64e-6 and rounds to a non-zero EMU, so it stays legal.
    const tiny = 1.5 / EMU_PER_INCH;
    expect(tiny).toBeLessThan(2e-6);
    expect(Math.round(tiny * EMU_PER_INCH)).toBe(2);
    expect(() => validatePictureGeometry({ width: tiny, height: tiny })).not.toThrow();
    const xml = JSON.stringify(pictureXml('2', 'rId2', { x: -1, y: -2.5, width: tiny, height: tiny }));
    expect(xml).toContain('"@_cx":"2"');
    expect(xml).toContain('"@_cy":"2"');
    expect(xml).toContain('"@_x":"-914400"');
    expect(xml).toContain('"@_y":"-2286000"');
  });

  it('keeps the defaults (0.5" position, natural image size) unchanged', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const info = slide.addPicture(png(96, 48));
    expect(info).toMatchObject({ x: 0.5, y: 0.5, width: 1, height: 0.5 });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const slideXml = (await zip.file('ppt/slides/slide1.xml')!.async('string'));
    expect(slideXml).toContain('<a:off x="457200" y="457200"/>');
    expect(slideXml).toContain('<a:ext cx="914400" cy="457200"/>');
    const tiny = slide.addPicture(png(96, 48), { width: ZERO_EMU_LIMIT, height: 1 / EMU_PER_INCH });
    expect(tiny).toMatchObject({ width: 1 / EMU_PER_INCH, height: 1 / EMU_PER_INCH });
  });
});
