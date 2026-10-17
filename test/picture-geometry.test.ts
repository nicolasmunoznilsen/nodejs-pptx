import JSZip from 'jszip';
import { deflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/index.js';
import { validatePictureGeometry } from '../src/shapes/picture.js';

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

describe('issue #37: picture position and size validation', () => {
  it('accepts absent options, zero and negative positions, and any finite size', () => {
    expect(() => validatePictureGeometry()).not.toThrow();
    expect(() => validatePictureGeometry({})).not.toThrow();
    expect(() => validatePictureGeometry({ name: 'Logo' })).not.toThrow();
    expect(() => validatePictureGeometry({ x: undefined, y: undefined, width: undefined, height: undefined })).not.toThrow();
    expect(() => validatePictureGeometry({ x: 0, y: 0, width: 0.001, height: 1e6 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -12.5, y: -0.25 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -100, y: 400, width: 2, height: 3 })).not.toThrow();
  });

  it('rejects non-number or non-finite positions naming the option and its value', () => {
    expect(() => validatePictureGeometry({ x: '1' as unknown as number })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ x: '1' as unknown as number })).toThrow(/^Invalid picture x: expected a finite number, got "1"\./);
    expect(() => validatePictureGeometry({ y: NaN })).toThrow(/^Invalid picture y: expected a finite number, got NaN\./);
    expect(() => validatePictureGeometry({ x: Infinity })).toThrow(/^Invalid picture x: expected a finite number, got Infinity\./);
    expect(() => validatePictureGeometry({ y: -Infinity })).toThrow(/^Invalid picture y: expected a finite number, got -Infinity\./);
    expect(() => validatePictureGeometry({ x: true as unknown as boolean })).toThrow(/^Invalid picture x: expected a finite number, got true\./);
  });

  it('rejects non-positive or non-finite dimensions naming the option and its value', () => {
    expect(() => validatePictureGeometry({ width: 0 })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ width: 0 })).toThrow(/^Invalid picture width: expected a finite number greater than 0, got 0\./);
    expect(() => validatePictureGeometry({ height: 0 })).toThrow(/^Invalid picture height: expected a finite number greater than 0, got 0\./);
    expect(() => validatePictureGeometry({ width: -3 })).toThrow(/^Invalid picture width: expected a finite number greater than 0, got -3\./);
    expect(() => validatePictureGeometry({ height: NaN })).toThrow(/^Invalid picture height: expected a finite number greater than 0, got NaN\./);
    expect(() => validatePictureGeometry({ width: Infinity })).toThrow(/^Invalid picture width: expected a finite number greater than 0, got Infinity\./);
    expect(() => validatePictureGeometry({ height: '2' as unknown as number })).toThrow(/^Invalid picture height: expected a finite number greater than 0, got "2"\./);
  });

  it('keeps the current defaults when geometry options are absent', async () => {
    const slide = Presentation.create().slides.add();
    const info = slide.addPicture(png(96, 48));
    expect(info.x).toBe(0.5);
    expect(info.y).toBe(0.5);
    expect(info.width).toBeCloseTo(1, 6);
    expect(info.height).toBeCloseTo(0.5, 6);
  });

  it('accepts off-canvas positions and stores them as EMU values', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addPicture(png(96, 48), { x: -2, y: 12.75, width: 3, height: 1.5 });
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).toContain('<a:off x="-1828800" y="11658600"/>');
    expect(slideXml).toContain('<a:ext cx="2743200" cy="1371600"/>');
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.get(0).getPictures()[0]).toMatchObject({ x: -2, y: 12.75, width: 3, height: 1.5 });
  });

  it('leaves media, rels, [Content_Types].xml and spTree untouched when geometry is invalid', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('keep me', { name: 'Caption' });
    const beforeZip = await JSZip.loadAsync(await pptx.toBuffer());
    const beforeNames = Object.keys(beforeZip.files).sort();

    const invalid: Array<Record<string, unknown>> = [
      { width: 0 },
      { height: 0 },
      { width: -1 },
      { height: NaN },
      { x: Infinity },
      { y: -Infinity },
      { y: '2' },
      { x: true },
    ];
    for (const options of invalid) {
      expect(() => slide.addPicture(png(96, 48), options as never), JSON.stringify(options)).toThrow(TypeError);
    }

    expect(slide.getPictures()).toHaveLength(0);
    expect(slide.shapes.all.map(shape => shape.type)).toEqual(['textBox']);

    const after = await JSZip.loadAsync(await pptx.toBuffer());
    expect(Object.keys(after.files).sort()).toEqual(beforeNames);
    expect(Object.keys(after.files).filter(name => name.startsWith('ppt/media/'))).toEqual([]);
    const types = await after.file('[Content_Types].xml')!.async('string');
    expect(types).not.toContain('image/png');
    expect(types).not.toContain('image/jpeg');
    const rels = await after.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
    expect(rels).not.toContain('/image');
    const slideXml = await after.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).not.toContain('<p:pic>');
    expect(slideXml).toContain('keep me');
  });

  it('validates geometry before the image payload is inspected', () => {
    const slide = Presentation.create().slides.add();
    expect(() => slide.addPicture(Buffer.from('GIF89a\x01\x00'), { width: 0 })).toThrow(/^Invalid picture width: /);
    expect(() => slide.addPicture(Buffer.alloc(0), { x: NaN })).toThrow(/^Invalid picture x: /);
  });
});

describe('issue #39: reject null and non-representable EMU geometry', () => {
  it('rejects null naming the option and rendering the value as null', () => {
    for (const name of ['x', 'y', 'width', 'height'] as const) {
      expect(() => validatePictureGeometry({ [name]: null } as unknown as Parameters<typeof validatePictureGeometry>[0])).toThrow(TypeError);
      expect(() => validatePictureGeometry({ [name]: null } as unknown as Parameters<typeof validatePictureGeometry>[0]))
        .toThrow(new RegExp(`^Invalid picture ${name}: expected a finite number( greater than 0)?, got null\\.`));
    }
  });

  it('rejects values whose EMU equivalent is not a safe integer', () => {
    expect(() => validatePictureGeometry({ x: Number.MAX_SAFE_INTEGER })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ x: 1e13 })).toThrow(/^Invalid picture x: expected a finite number whose EMU value is a safe integer, got 10000000000000\./);
    expect(() => validatePictureGeometry({ y: -1e13 })).toThrow(/^Invalid picture y: expected a finite number whose EMU value is a safe integer, got -10000000000000\./);
    expect(() => validatePictureGeometry({ width: 1e13 })).toThrow(/^Invalid picture width: expected a finite number greater than 0 whose EMU value is a safe integer, got 10000000000000\./);
    expect(() => validatePictureGeometry({ height: 1e10 })).toThrow(/^Invalid picture height: expected a finite number greater than 0 whose EMU value is a safe integer, got 10000000000\./);
  });

  it('keeps accepting representable geometry, including the defaults and off-canvas values', () => {
    expect(() => validatePictureGeometry({ x: undefined, y: undefined, width: undefined, height: undefined })).not.toThrow();
    expect(() => validatePictureGeometry({ x: 0, y: 0, width: 0.001, height: 1e6 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -12.5, y: -0.25 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -100, y: 400, width: 2, height: 3 })).not.toThrow();
    expect(() => validatePictureGeometry({ name: 'Logo' })).not.toThrow();
  });

  it('keeps the defaults when geometry options are absent', async () => {
    const info = Presentation.create().slides.add().addPicture(png(96, 48));
    expect(info.x).toBe(0.5);
    expect(info.y).toBe(0.5);
    expect(info.width).toBeCloseTo(1, 6);
    expect(info.height).toBeCloseTo(0.5, 6);
  });

  it('rejects null and oversized geometry before any package mutation', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('keep me', { name: 'Caption' });
    const beforeNames = Object.keys((await JSZip.loadAsync(await pptx.toBuffer())).files).sort();

    const invalid: Array<Record<string, unknown>> = [
      { x: null },
      { y: null },
      { width: null },
      { height: null },
      { x: 1e13 },
      { height: 1e10 },
    ];
    for (const options of invalid) {
      expect(() => slide.addPicture(png(96, 48), options as never)).toThrow(/^Invalid picture (x|y|width|height): /);
    }

    const after = await JSZip.loadAsync(await pptx.toBuffer());
    expect(Object.keys(after.files).sort()).toEqual(beforeNames);
    const slideXml = await after.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).not.toContain('<p:pic>');
    expect(slideXml).toContain('keep me');
  });

  it('validates geometry before the image payload is inspected', () => {
    const slide = Presentation.create().slides.add();
    expect(() => slide.addPicture(Buffer.alloc(0), { x: null as unknown as number })).toThrow(/^Invalid picture x: expected a finite number, got null\./);
    expect(() => slide.addPicture(Buffer.alloc(0), { width: 1e13 })).toThrow(/^Invalid picture width: expected a finite number greater than 0 whose EMU value is a safe integer/);
  });
});
