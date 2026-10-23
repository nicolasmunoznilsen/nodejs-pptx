import { EMU_PER_INCH } from '../ooxml/constants.js';
import type { XmlNode } from '../ooxml/xml.js';

function inches(value: number): number { return Math.round(value * EMU_PER_INCH); }

/** EMU per pixel at the 96 DPI reference used by OOXML drawings. */
const EMU_PER_PIXEL = EMU_PER_INCH / 96;
const DEFAULT_PIXEL_WIDTH = 640;
const DEFAULT_PIXEL_HEIGHT = 480;

export type ImageFormat = 'png' | 'jpeg';

export interface PictureOptions { x?: number; y?: number; width?: number; height?: number; name?: string }
export interface PictureInfo { name: string; id: string; index: number; x: number; y: number; width: number; height: number; fileName: string; contentType: string; format?: ImageFormat; relationshipId: string }

export interface ImageData {
  format: ImageFormat;
  /** Package extension used for the media part, without a leading dot. */
  extension: string;
  contentType: string;
  /** Extensions that need a content-type Default entry for this image. */
  defaultExtensions: string[];
  buffer: Buffer | Uint8Array;
  /** Pixel dimensions when they could be read from the image header. */
  width?: number;
  height?: number;
}

const IMAGE_TYPES: Record<ImageFormat, { extension: string; contentType: string; defaultExtensions: string[] }> = {
  png: { extension: 'png', contentType: 'image/png', defaultExtensions: ['png'] },
  jpeg: { extension: 'jpg', contentType: 'image/jpeg', defaultExtensions: ['jpg', 'jpeg'] },
};

export function contentTypeForExtension(extension: string): string | undefined {
  const ext = extension.replace(/^\.+/, '').toLowerCase();
  if (ext === 'png') return 'image/png';
  if (ext === 'jpg' || ext === 'jpeg' || ext === 'jpe') return 'image/jpeg';
  if (ext === 'gif') return 'image/gif';
  if (ext === 'bmp') return 'image/bmp';
  if (ext === 'tiff' || ext === 'tif') return 'image/tiff';
  return undefined;
}

/** Detects an image from its magic bytes and reads its pixel size when available. */
export function detectImage(buffer: Buffer | Uint8Array): ImageData {
  const bytes = asBytes(buffer);
  if (!bytes || bytes.length < 4) throw new Error(`Unsupported image format: ${nameFormat(bytes)}. Only PNG and JPEG images can be inserted.`);
  if (isPng(bytes)) {
    if (!isCompletePng(bytes)) throw new Error(`Unsupported image format: ${nameFormat(bytes)}. Only PNG and JPEG images can be inserted.`);
    return { format: 'png', ...IMAGE_TYPES.png, buffer: bytes, ...pngSize(bytes) };
  }
  if (isJpeg(bytes)) {
    if (!hasJpegEoi(bytes)) throw new Error(`Unsupported image format: ${nameFormat(bytes)}. Only PNG and JPEG images can be inserted.`);
    return { format: 'jpeg', ...IMAGE_TYPES.jpeg, buffer: bytes, ...jpegSize(bytes) };
  }
  throw new Error(`Unsupported image format: ${nameFormat(bytes)}. Only PNG and JPEG images can be inserted.`);
}

function asBytes(input: Buffer | Uint8Array | undefined): Uint8Array | undefined {
  if (!input) return undefined;
  return input instanceof Uint8Array ? input : new Uint8Array(input as Uint8Array);
}

/** Names a well-known image format so unsupported input produces an actionable error. */
function nameFormat(b: Uint8Array | undefined): string {
  if (!b || b.length === 0) return 'empty or missing image data';
  const text = ascii(b, 0, Math.min(b.length, 8));
  if (hasPngSignature(b)) return 'truncated PNG';
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xd8) return 'truncated JPEG';
  if (text.startsWith('GIF8')) return 'GIF';
  if (text.startsWith('BM')) return 'BMP';
  if (text.startsWith('II') || text.startsWith('MM')) return 'TIFF';
  if (text.startsWith('RIFF') && ascii(b, 8, 12) === 'WEBP') return 'WEBP';
  if (text.startsWith('<svg') || text.startsWith('<?xm')) return 'SVG';
  if (b.length >= 4 && le32(b, 0) === 1) return 'EMF';
  if (b.length >= 4 && (le32(b, 0) === 0xd7cdc69a || u32(b, 0) === 0xd7cdc69a)) return 'WMF';
  if (isWmfMetaheader(b)) return 'WMF';
  return 'unknown';
}

/**
 * True for a standard (non-placeable) WMF `METAHEADER`, the half of the WMF magic
 * the placeable signature above cannot see: the type and delay words at bytes 0-3
 * are zero, bytes 4-5 hold a plausible header size in words, the planes and object
 * counts are in range, and a known `nversion` word sits at offset 18. Pure, and
 * safe on short buffers because every read sits behind the length guard.
 */
function isWmfMetaheader(b: Uint8Array | undefined): boolean {
  if (!b || b.length < 22) return false;
  if (le32(b, 0) !== 0) return false;
  const headerSize = le16(b, 4);
  if (headerSize < 9 || headerSize > 0x154) return false;
  if (le16(b, 6) > 2 || le16(b, 8) >= 0x8000) return false;
  return WMF_VERSIONS.includes(le16(b, 18)) || WMF_VERSIONS.includes(u16(b, 18));
}

/** `nversion` values published by the 1.x-3.x WMF releases. */
const WMF_VERSIONS = [0x0100, 0x0200, 0x0237, 0x0300, 0x0310];

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function isPng(b: Uint8Array): boolean { return b.length >= 8 && hasPngSignature(b); }

/** Matches as much of the eight-byte PNG signature as the payload still carries. */
function hasPngSignature(b: Uint8Array): boolean {
  if (b.length < 4) return false;
  for (let i = 0; i < Math.min(b.length, PNG_SIGNATURE.length); i += 1) {
    if (b[i] !== PNG_SIGNATURE[i]) return false;
  }
  return true;
}

/** Walks the chunk chain so a payload cut short of its `IEND` chunk is rejected. */
function isCompletePng(b: Uint8Array): boolean {
  if (b.length < 24 || u32(b, 8) < 13 || ascii(b, 12, 16) !== 'IHDR') return false;
  let offset = 8;
  while (offset + 12 <= b.length) {
    const length = u32(b, offset);
    const type = ascii(b, offset + 4, offset + 8);
    if (offset + 12 + length > b.length) return false;
    if (type === 'IEND') return length === 0;
    offset += 12 + length;
  }
  return false;
}

function isJpeg(b: Uint8Array): boolean { return b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff; }

/** A JPEG is only complete once its `EOI` marker is present. */
function hasJpegEoi(b: Uint8Array): boolean {
  for (let offset = 2; offset + 1 < b.length; offset += 1) {
    if (b[offset] === 0xff && b[offset + 1] === 0xd9) return true;
  }
  return false;
}

function u16(b: Uint8Array, offset: number): number { return (b[offset] << 8) | b[offset + 1]; }
function le16(b: Uint8Array, offset: number): number { return b[offset] | (b[offset + 1] << 8); }
function u32(b: Uint8Array, offset: number): number { return ((b[offset] * 16777216) + (b[offset + 1] << 16) + (b[offset + 2] << 8)) + b[offset + 3]; }
function le32(b: Uint8Array, offset: number): number { return b[offset] + (b[offset + 1] << 8) + (b[offset + 2] << 16) + (b[offset + 3] * 16777216); }
function ascii(b: Uint8Array, offset: number, end: number): string { return String.fromCharCode.apply(null, Array.prototype.slice.call(b, offset, end) as number[]); }

/** PNG IHDR stores the pixel size as big-endian uint32 values at bytes 16-23. */
function pngSize(b: Uint8Array): { width?: number; height?: number } {
  if (b.length < 24 || u32(b, 8) < 13 || ascii(b, 12, 16) !== 'IHDR') return {};
  const width = u32(b, 16);
  const height = u32(b, 20);
  return width > 0 && height > 0 ? { width, height } : {};
}

/** Walks JPEG markers looking for an SOFn frame that carries the pixel size. */
function jpegSize(b: Uint8Array): { width?: number; height?: number } {
  let offset = 2;
  while (offset + 9 < b.length) {
    if (b[offset] !== 0xff) { offset += 1; continue; }
    const marker = b[offset + 1];
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { offset += 2; continue; }
    if (marker === 0xda || marker === 0xd9) break;
    const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    const length = u16(b, offset + 2);
    if (length < 2) break;
    if (isSof) {
      const height = u16(b, offset + 5);
      const width = u16(b, offset + 7);
      if (width > 0 && height > 0) return { width, height };
    }
    offset += 2 + length;
  }
  return {};
}

/**
 * Natural display size, in inches, of an image at 96 DPI. The pixel dimensions come from
 * {@link detectImage}, which only reports a size when it is greater than 0, so every result
 * is at least one pixel (1/96\" ≈ 9525 EMU) or falls back to the 640x480 default: neither can
 * ever round to 0 EMU, so a natural size always satisfies {@link validatePictureGeometry}.
 */
export function naturalImageSize(image: ImageData): { width: number; height: number } {
  const width = (image.width ?? DEFAULT_PIXEL_WIDTH) * EMU_PER_PIXEL / EMU_PER_INCH;
  const height = (image.height ?? DEFAULT_PIXEL_HEIGHT) * EMU_PER_PIXEL / EMU_PER_INCH;
  return { width, height };
}

/** @internal Shared message for rejected picture geometry options. */
function describeInvalidGeometry(name: string, value: unknown, positive: boolean, representable = false): string {
  const expected = representable
    ? `a finite number${positive ? ' greater than 0' : ''} whose EMU value is a safe integer`
    : positive ? 'a finite number greater than 0' : 'a finite number';
  return `Invalid picture ${name}: expected ${expected}, got ${describeGeometryValue(value)}.`;
}

/** @internal Renders a rejected option value the way the error message should report it. */
function describeGeometryValue(value: unknown): string {
  if (typeof value === 'string') return JSON.stringify(value);
  if (typeof value === 'number') return String(value);
  if (typeof value === 'bigint' || typeof value === 'boolean' || value === null || value === undefined) return String(value);
  return JSON.stringify(value) ?? String(value);
}

/**
 * Validates the position and size options of a `PictureOptions` object before any package
 * mutation happens. Every option that is present must be a `number` and `Number.isFinite`;
 * `width` and `height` must additionally be greater than 0, and every value must convert to
 * a `Number.isSafeInteger` EMU so it can be written to the OOXML geometry. `width` and
 * `height` must also convert to a non-zero EMU: a positive size that rounds to `0` EMU
 * (`|value| < 0.5 / EMU_PER_INCH`, roughly `5.4677e-7` inches, so `Number.MIN_VALUE`,
 * `5e-7` and `1e-9` all land here) describes an invisible picture that PowerPoint cannot
 * render, so it is rejected with the same message as the other non-representable EMU
 * values. `x` and `y` accept any finite number, including negative values and values that
 * round to `0` EMU, because a picture outside the canvas or flush with the origin is legal
 * OOXML. Only absent (`undefined`) options are left untouched so the existing defaults
 * apply (0.5\" for position, the natural image size for dimensions); `null` is rejected like
 * any other invalid value.
 *
 * Throws a `TypeError` naming the offending option and its value.
 */
export function validatePictureGeometry(options: PictureOptions = {}): void {
  for (const name of ['x', 'y', 'width', 'height'] as const) {
    const value = options[name];
    if (value === undefined) continue;
    const positive = name === 'width' || name === 'height';
    if (typeof value !== 'number' || !Number.isFinite(value) || (positive && value <= 0)) {
      throw new TypeError(describeInvalidGeometry(name, value, positive));
    }
    const emu = Math.round(value * EMU_PER_INCH);
    if (!Number.isSafeInteger(emu) || (positive && emu === 0)) {
      throw new TypeError(describeInvalidGeometry(name, value, positive, true));
    }
  }
}

/**
 * Builds the `p:pic` element. Both the caller-supplied options and the effective size (the
 * `size` fallback that carries the natural image dimensions) go through
 * {@link validatePictureGeometry}, so no geometry that cannot be represented in EMU reaches
 * the package, whether it came from an option or from a default.
 */
export function pictureXml(id: string, rId: string, options: PictureOptions = {}, size = { width: 4, height: 3 }): XmlNode {
  validatePictureGeometry(options);
  validatePictureGeometry({ width: options.width ?? size.width, height: options.height ?? size.height });
  return {
    'p:nvPicPr': {
      'p:cNvPr': { '@_id': id, '@_name': options.name ?? `Picture ${id}` },
      'p:cNvPicPr': { 'a:picLocks': { '@_noChangeAspect': '1' } },
      'p:nvPr': {},
    },
    'p:blipFill': { 'a:blip': { '@_r:embed': rId }, 'a:stretch': { 'a:fillRect': {} } },
    'p:spPr': {
      'a:xfrm': {
        'a:off': { '@_x': String(inches(options.x ?? 0.5)), '@_y': String(inches(options.y ?? 0.5)) },
        'a:ext': { '@_cx': String(inches(options.width ?? size.width)), '@_cy': String(inches(options.height ?? size.height)) },
      },
      'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} },
    },
  };
}
