import { readFileSync } from 'node:fs';
import type JSZip from 'jszip';
import { allocateMediaPath, nextShapeId } from '../ooxml/ids.js';
import { IMAGE_REL } from '../ooxml/constants.js';
import { relsPath, relativePackageTarget, resolveRelationshipTarget } from '../ooxml/paths.js';
import { arr, parseXml, readAttr, walk, xmlToString, type XmlNode } from '../ooxml/xml.js';
import type { ContentTypes } from '../package/content-types.js';
import { RelationshipCollection } from '../package/relationships.js';
import { ShapeCollection } from '../shapes/shape-collection.js';
import { detectImage, naturalImageSize, pictureXml, type ImageData, type PictureInfo, type PictureOptions } from '../shapes/picture.js';
import { textBoxXml, type TextBoxInfo, type TextBoxOptions } from '../shapes/text-box.js';
import type { Shape } from '../shapes/shape.js';
import { PlaceholderCollection, placeholdersFromTree } from './placeholders.js';
import type { SlideLayout } from '../presentation/slide-layout.js';

export class Slide {
  dirty = false;
  relsDirty = false;
  private readonly relationships: RelationshipCollection;
  readonly shapes: ShapeCollection;
  layout?: SlideLayout;

  constructor(private readonly zip: JSZip, readonly path: string, private xml: XmlNode, relsXml?: XmlNode, private readonly contentTypes?: ContentTypes) {
    this.relationships = new RelationshipCollection(relsXml);
    this.shapes = new ShapeCollection(this.spTree, () => this.markDirty(), rid => this.resolveImageTarget(rid));
  }

  get placeholders(): PlaceholderCollection { return new PlaceholderCollection(placeholdersFromTree(this.spTree, () => this.markDirty())); }

  get background(): SlideBackground { return new SlideBackground(this.xml, () => this.markDirty()); }

  get text(): string { return this.shapes.all.filter(s => s.hasTextFrame).map(s => s.textFrame.text).filter(Boolean).join('\n'); }

  getTextBoxes(): TextBoxInfo[] {
    return this.shapes.all.filter(s => s.hasTextFrame && (s.type === 'textBox' || s.type === 'autoShape')).map((shape, index) => ({
      text: shape.textFrame.text,
      name: shape.name,
      id: shape.id,
      index,
      x: shape.x,
      y: shape.y,
      width: shape.width,
      height: shape.height,
      alignment: shape.textFrame.alignment,
      margin: shape.textFrame.margin,
      marginLeft: shape.textFrame.marginLeft,
      marginTop: shape.textFrame.marginTop,
      marginRight: shape.textFrame.marginRight,
      marginBottom: shape.textFrame.marginBottom,
    }));
  }

  replaceText(search: string | RegExp, replacement: string): number {
    let count = 0;
    for (const shape of this.shapes.all) {
      if (!shape.hasTextFrame) continue;
      for (const paragraph of shape.textFrame.paragraphs) {
        count += replaceInRuns(paragraph.runs, search, replacement);
      }
    }
    if (count) this.markDirty();
    return count;
  }

  addTextBox(text: string, options: TextBoxOptions = {}): TextBoxInfo {
    const shapes = arr(this.spTree['p:sp'] as XmlNode | XmlNode[] | undefined);
    const id = nextShapeId(this.spTree);
    shapes.push(textBoxXml(text, id, options));
    this.spTree['p:sp'] = shapes;
    this.persist();
    return this.getTextBoxes().at(-1)!;
  }

  getPictures(): PictureInfo[] {
    return this.shapes.all.filter(shape => shape.type === 'picture').map((shape, index) => ({
      name: shape.name,
      id: shape.id,
      index,
      x: shape.x,
      y: shape.y,
      width: shape.width,
      height: shape.height,
      fileName: shape.imageFileName ?? '',
      contentType: shape.imageContentType ?? '',
      format: shape.imageFormat,
      relationshipId: shape.imageRelationshipId ?? '',
    }));
  }

  /** Inserts a PNG or JPEG image, given as a Buffer, a Uint8Array, or a filesystem path. */
  addPicture(input: Buffer | Uint8Array | string, options: PictureOptions = {}): PictureInfo {
    const image = detectImage(typeof input === 'string' ? readFileSync(input) : input);
    const mediaPath = allocateMediaPath(this.zip, image.extension);
    this.zip.file(mediaPath, image.buffer);
    image.defaultExtensions.forEach(ext => this.contentTypes?.addDefault(ext, image.contentType));
    const rel = this.relationships.add(IMAGE_REL, relativePackageTarget(this.path, mediaPath));
    this.relsDirty = true;
    const pictures = arr(this.spTree['p:pic'] as XmlNode | XmlNode[] | undefined);
    pictures.push(pictureXml(nextShapeId(this.spTree), rel.id, options, naturalImageSize(image)));
    this.spTree['p:pic'] = pictures;
    this.persist();
    this.zip.file(relsPath(this.path), xmlToString(this.relationships.xml));
    return this.getPictures().at(-1)!;
  }

  /** @internal */ private resolveImageTarget(relationshipId: string): string | undefined {
    const rel = this.relationships.getById(relationshipId);
    if (!rel || rel.type !== IMAGE_REL || !rel.target) return undefined;
    return resolveRelationshipTarget(this.path, rel.target);
  }

  persist(): void { this.markDirty(); this.zip.file(this.path, xmlToString(this.xml)); }
  persistIfDirty(): void {
    if (this.dirty) this.zip.file(this.path, xmlToString(this.xml));
    if (this.relationships.dirty || this.relsDirty) this.zip.file(relsPath(this.path), xmlToString(this.relationships.xml));
  }

  /** @internal */ get xmlNode(): XmlNode { return this.xml; }
  /** @internal */ get relationshipsXml(): XmlNode { return this.relationships.xml; }
  /** @internal */ get relationshipsItems() { return this.relationships.items; }
  private markDirty(): void { this.dirty = true; }
  private get spTree(): XmlNode {
    const tree = (((this.xml['p:sld'] as XmlNode)['p:cSld'] as XmlNode)['p:spTree']) as XmlNode | undefined;
    if (!tree) throw new Error('Invalid slide XML: missing p:spTree.');
    return tree;
  }

  static parse(zip: JSZip, path: string, xmlString: string, relsString?: string, contentTypes?: ContentTypes): Slide {
    return new Slide(zip, path, parseXml(path, xmlString), relsString ? parseXml(relsPath(path), relsString) : undefined, contentTypes);
  }
}

function replaceInRuns(runs: Array<{ text: string }>, search: string | RegExp, replacement: string): number {
  const full = runs.map(r => r.text).join('');
  const matches = typeof search === 'string' ? findStringMatches(full, search) : findRegexMatches(full, search);
  for (const match of matches.reverse()) applyReplacement(runs, match.start, match.end, replacement);
  return matches.length;
}

function findStringMatches(text: string, needle: string): Array<{ start: number; end: number }> {
  if (!needle) return [];
  const matches: Array<{ start: number; end: number }> = [];
  let start = text.indexOf(needle);
  while (start >= 0) {
    matches.push({ start, end: start + needle.length });
    start = text.indexOf(needle, start + needle.length);
  }
  return matches;
}

function findRegexMatches(text: string, regex: RegExp): Array<{ start: number; end: number }> {
  const flags = regex.flags.includes('g') ? regex.flags : `${regex.flags}g`;
  const clone = new RegExp(regex.source, flags);
  const matches: Array<{ start: number; end: number }> = [];
  for (const match of text.matchAll(clone)) {
    if (match.index === undefined || match[0] === '') continue;
    matches.push({ start: match.index, end: match.index + match[0].length });
  }
  return matches;
}

function applyReplacement(runs: Array<{ text: string }>, start: number, end: number, replacement: string): void {
  let pos = 0;
  let wrote = false;
  for (const run of runs) {
    const runStart = pos;
    const runEnd = pos + run.text.length;
    pos = runEnd;
    if (runEnd <= start || runStart >= end) continue;
    const before = run.text.slice(0, Math.max(0, start - runStart));
    const after = run.text.slice(Math.max(0, end - runStart));
    if (!wrote) { run.text = before + replacement + after; wrote = true; } else { run.text = after; }
  }
}

export function blankSlideXml(): string { return `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`; }

export class SlideBackground {
  constructor(private readonly xml: XmlNode, private readonly markDirty: () => void) {}
  get followMasterBackground(): boolean { return !this.bg; }
  set followMasterBackground(value: boolean) { if (value && this.bg) { delete this.cSld['p:bg']; this.markDirty(); } }
  get color(): string | undefined { return readAttr((((this.bg?.['p:bgPr'] as XmlNode|undefined)?.['a:solidFill'] as XmlNode|undefined)?.['a:srgbClr'] as XmlNode|undefined), 'val'); }
  set color(value: string | undefined) {
    if (!value) { if (this.bg) delete this.cSld['p:bg']; this.markDirty(); return; }
    this.cSld['p:bg'] = { 'p:bgPr': { 'a:solidFill': { 'a:srgbClr': { '@_val': value.replace(/^#/, '') } } } };
    this.markDirty();
  }
  private get root(): XmlNode { return this.xml['p:sld'] as XmlNode; }
  private get cSld(): XmlNode { return this.root['p:cSld'] as XmlNode; }
  private get bg(): XmlNode | undefined { return this.cSld['p:bg'] as XmlNode | undefined; }
}
