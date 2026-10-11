import type JSZip from 'jszip';
import { nextShapeId } from '../ooxml/ids.js';
import { relsPath } from '../ooxml/paths.js';
import { arr, parseXml, readAttr, walk, xmlToString, type XmlNode } from '../ooxml/xml.js';
import { RelationshipCollection } from '../package/relationships.js';
import { ShapeCollection } from '../shapes/shape-collection.js';
import { textBoxXml, type TextBoxInfo, type TextBoxOptions } from '../shapes/text-box.js';
import type { Shape } from '../shapes/shape.js';
import { PlaceholderCollection, placeholdersFromTree } from './placeholders.js';
import { replaceInRuns } from '../text/replace.js';
import type { SlideLayout } from '../presentation/slide-layout.js';

export class Slide {
  dirty = false;
  relsDirty = false;
  private readonly relationships: RelationshipCollection;
  readonly shapes: ShapeCollection;
  layout?: SlideLayout;

  constructor(private readonly zip: JSZip, readonly path: string, private xml: XmlNode, relsXml?: XmlNode) {
    this.relationships = new RelationshipCollection(relsXml);
    this.shapes = new ShapeCollection(this.spTree, () => this.markDirty());
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
    }));
  }

  /**
   * Replace every occurrence of `search` in this slide, including matches that cross run
   * boundaries inside a paragraph, and return how many replacements were made. Placeholder shapes
   * and shapes nested inside groups are covered; only `a:t` text nodes are touched.
   */
  replaceText(search: string | RegExp, replacement: string): number {
    let count = 0;
    for (const shape of this.shapes.allDescendants) {
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

  static parse(zip: JSZip, path: string, xmlString: string, relsString?: string): Slide {
    return new Slide(zip, path, parseXml(path, xmlString), relsString ? parseXml(relsPath(path), relsString) : undefined);
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
