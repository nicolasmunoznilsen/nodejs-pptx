import JSZip from 'jszip';
import { deflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/index.js';
import { validatePictureGeometry } from '../src/shapes/picture.js';
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

/** Smallest inch value whose rounded EMU is no longer a safe integer. */
const UNSAFE_INCHES = (Number.MAX_SAFE_INTEGER + 1) / EMU_PER_INCH;

describe('issue #39: null geometry is rejected like any other invalid value', () => {
  it('rejects a null x or y with the existing finite-number message', () => {
    expect(() => validatePictureGeometry({ x: null as unknown as number })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ x: null as unknown as number }))
      .toThrow(/^Invalid picture x: expected a finite number, got null\./);
    expect(() => validatePictureGeometry({ y: null as unknown as number })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ y: null as unknown as number }))
      .toThrow(/^Invalid picture y: expected a finite number, got null\./);
  });

  it('rejects a null width or height with the existing positive-number message', () => {
    expect(() => validatePictureGeometry({ width: null as unknown as number })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ width: null as unknown as number }))
      .toThrow(/^Invalid picture width: expected a finite number greater than 0, got null\./);
    expect(() => validatePictureGeometry({ height: null as unknown as number })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ height: null as unknown as number }))
      .toThrow(/^Invalid picture height: expected a finite number greater than 0, got null\./);
  });

  it('still accepts absent (undefined) geometry for all four options', () => {
    expect(() => validatePictureGeometry()).not.toThrow();
    expect(() => validatePictureGeometry({})).not.toThrow();
    expect(() => validatePictureGeometry({
      x: undefined, y: undefined, width: undefined, height: undefined,
    })).not.toThrow();
    expect(() => validatePictureGeometry({
      x: undefined, y: undefined, width: undefined, height: undefined, name: 'Logo',
    })).not.toThrow();
  });
});

describe('issue #39: non-representable EMU geometry is rejected', () => {
  it('rejects x and y whose rounded EMU exceeds Number.MAX_SAFE_INTEGER', () => {
    expect(Number.isSafeInteger(Math.round(1e13 * EMU_PER_INCH))).toBe(false);
    expect(() => validatePictureGeometry({ x: 1e13 })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ x: 1e13 })).toThrow(/^Invalid picture x: /);
    expect(() => validatePictureGeometry({ y: 1e13 })).toThrow(/^Invalid picture y: /);
    expect(() => validatePictureGeometry({ x: -1e13 })).toThrow(/^Invalid picture x: /);
    expect(() => validatePictureGeometry({ y: -1e13 })).toThrow(/^Invalid picture y: /);
    expect(() => validatePictureGeometry({ x: 1e300 })).toThrow(/^Invalid picture x: /);
    expect(() => validatePictureGeometry({ y: -1e300 })).toThrow(/^Invalid picture y: /);
    expect(() => validatePictureGeometry({ x: Number.MAX_VALUE })).toThrow(/^Invalid picture x: /);
  });

  it('rejects width and height whose rounded EMU exceeds Number.MAX_SAFE_INTEGER', () => {
    expect(() => validatePictureGeometry({ width: 1e13 })).toThrow(TypeError);
    expect(() => validatePictureGeometry({ width: 1e13 }))
      .toThrow(/^Invalid picture width: expected a finite number greater than 0.*got 10000000000000\./);
    expect(() => validatePictureGeometry({ height: 1e13 })).toThrow(/^Invalid picture height: .*got 10000000000000\./);
    expect(() => validatePictureGeometry({ width: 1e300 })).toThrow(/^Invalid picture width: /);
    expect(() => validatePictureGeometry({ height: 1e300 })).toThrow(/^Invalid picture height: /);
  });

  it('rejects exactly at the safe-integer EMU boundary and keeps large-but-safe values', () => {
    // Boundary documentation: 914400 EMU per inch means anything at or above
    // (MAX_SAFE_INTEGER + 1) / 914400 inches can no longer be written as EMU.
    expect(UNSAFE_INCHES).toBeGreaterThan(9.85e9);
    expect(UNSAFE_INCHES).toBeLessThan(9.86e9);
    expect(() => validatePictureGeometry({ x: UNSAFE_INCHES })).toThrow(/^Invalid picture x: /);
    // Large but representable values must keep working, exactly as before issue #39.
    expect(() => validatePictureGeometry({ width: 1e6 })).not.toThrow();
    expect(() => validatePictureGeometry({ height: 1e6 })).not.toThrow();
    expect(() => validatePictureGeometry({ width: 9_000_000, height: 9_000_000 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: 9_000_000, y: -9_000_000 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: 0, y: 0, width: 0.001, height: 1e6 })).not.toThrow();
    expect(() => validatePictureGeometry({ x: -12.5, y: -0.25 })).not.toThrow();
  });
});

describe('issue #39: addPicture rejects before touching the package', () => {
  it('leaves media, slide rels, [Content_Types].xml and spTree byte-identical', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('keep me', { name: 'Caption' });
    const before = await snapshot(await pptx.toBuffer());
    const baselineMedia = [...before.keys()].filter(name => name.startsWith('ppt/media/'));
    expect(baselineMedia).toEqual([]);

    const invalid: Array<Record<string, unknown>> = [
      { x: null },
      { y: null },
      { width: null },
      { height: null },
      { x: 1e300 },
      { y: 1e13 },
      { width: 1e300 },
      { width: 1e13 },
      { height: 1e13 },
    ];
    for (const options of invalid) {
      expect(() => slide.addPicture(png(96, 48), options as never), JSON.stringify(options)).toThrow(TypeError);
      expect(() => slide.addPicture(png(96, 48), options as never), JSON.stringify(options))
        .toThrow(/^Invalid picture (x|y|width|height): /);
    }

    const after = await snapshot(await pptx.toBuffer());
    expectIdentical(before, after);
    expect([...after.keys()].filter(name => name.startsWith('ppt/media/'))).toEqual([]);
    expect((after.get('[Content_Types].xml') ?? Buffer.alloc(0)).toString())
      .not.toMatch(/image\/(png|jpeg)/);
    expect((after.get('ppt/slides/_rels/slide1.xml.rels') ?? Buffer.alloc(0)).toString())
      .not.toContain('/image');
    const slideXml = (after.get('ppt/slides/slide1.xml') ?? Buffer.alloc(0)).toString();
    expect(slideXml).not.toContain('<p:pic>');
    expect(slideXml).toContain('keep me');
    expect(slide.getPictures()).toHaveLength(0);
    expect(slide.shapes.all.map(shape => shape.type)).toEqual(['textBox']);
  });

  it('still produces the same defaults and getPictures() round-trip afterwards', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    expect(() => slide.addPicture(png(96, 48), { x: null as unknown as number })).toThrow(/^Invalid picture x: /);
    expect(() => slide.addPicture(png(96, 48), { width: 1e300 })).toThrow(/^Invalid picture width: /);

    const info = slide.addPicture(png(96, 48));
    expect(info.x).toBe(0.5);
    expect(info.y).toBe(0.5);
    expect(info.width).toBeCloseTo(1, 6);
    expect(info.height).toBeCloseTo(0.5, 6);

    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).toContain('<a:off x="457200" y="457200"/>');
    expect(slideXml).toContain('<a:ext cx="914400" cy="457200"/>');

    const reopened = await Presentation.open(await pptx.toBuffer());
    const [picture] = reopened.slides.get(0).getPictures();
    expect(picture).toMatchObject({ x: 0.5, y: 0.5, width: 1, height: 0.5 });

    const sized = reopened.slides.get(0).addPicture(png(96, 48), { x: 2, y: 3, width: 1e6, height: 4 });
    expect(sized).toMatchObject({ x: 2, y: 3, width: 1e6, height: 4 });
    const reread = await Presentation.open(await reopened.toBuffer());
    expect(reread.slides.get(0).getPictures()[1]).toMatchObject({ x: 2, y: 3, width: 1e6, height: 4 });
  });

  it('validates geometry before the image payload is inspected', () => {
    const slide = Presentation.create().slides.add();
    expect(() => slide.addPicture(Buffer.alloc(0), { x: null as unknown as number }))
      .toThrow(/^Invalid picture x: expected a finite number, got null\./);
    expect(() => slide.addPicture(Buffer.alloc(0), { width: 1e300 }))
      .toThrow(/^Invalid picture width: expected a finite number greater than 0/);
  });
});
