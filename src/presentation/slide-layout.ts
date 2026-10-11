import type JSZip from 'jszip';
import { basename } from 'node:path';
import { SLIDE_MASTER_REL } from '../ooxml/constants.js';
import { relsPath, resolveRelationshipTarget } from '../ooxml/paths.js';
import { arr, parseXml, readAttr, xmlToString, type XmlNode } from '../ooxml/xml.js';
import { RelationshipCollection } from '../package/relationships.js';
import { ShapeCollection } from '../shapes/shape-collection.js';
import { PlaceholderCollection, placeholdersFromTree } from '../slides/placeholders.js';
import { replaceInRuns } from '../text/replace.js';

export class SlideLayout {
  readonly placeholders: PlaceholderCollection;
  readonly shapes: ShapeCollection;
  dirty = false;
  constructor(readonly path: string, readonly xml: XmlNode, readonly master?: { path: string; name?: string }, private readonly zip?: JSZip) {
    const tree = layoutTree(xml) ?? {};
    const markDirty = () => { this.dirty = true; };
    this.shapes = new ShapeCollection(tree, markDirty);
    this.placeholders = new PlaceholderCollection(placeholdersFromTree(tree, markDirty));
  }
  get name(): string { return readAttr((this.xml['p:sldLayout'] as XmlNode)?.['p:cSld'], 'name') ?? basename(this.path, '.xml'); }
  get type(): string { return readAttr(this.xml['p:sldLayout'], 'type') ?? 'unknown'; }

  /**
   * Same cross-run replacement as `slide.replaceText()`, applied to the text of this layout. Layout
   * text is template content: it is normally inherited by every slide using the layout, so prefer
   * editing slides unless the template itself must change.
   */
  replaceText(search: string | RegExp, replacement: string): number {
    let count = 0;
    for (const shape of this.shapes.allDescendants) {
      if (!shape.hasTextFrame) continue;
      for (const paragraph of shape.textFrame.paragraphs) {
        count += replaceInRuns(paragraph.runs, search, replacement);
      }
    }
    if (count) this.dirty = true;
    return count;
  }

  /** @internal Write the layout part back to the package only when it was edited. */
  persistIfDirty(): void {
    if (!this.dirty || !this.zip) return;
    this.zip.file(this.path, xmlToString(this.xml));
    this.dirty = false;
  }
}

export class SlideLayoutCollection {
  constructor(private readonly layouts: SlideLayout[]) {}
  get length(): number { return this.layouts.length; }
  get(index: number): SlideLayout {
    if (!Number.isInteger(index) || index < 0 || index >= this.length) throw new RangeError(`Slide layout index ${index} is out of range (0-${this.length - 1}).`);
    return this.layouts[index];
  }
  findByName(name: string): SlideLayout | undefined { return this.layouts.find(l => l.name === name); }
  toArray(): SlideLayout[] { return [...this.layouts]; }
  [Symbol.iterator](): IterableIterator<SlideLayout> { return this.layouts[Symbol.iterator](); }
}

export async function loadSlideLayouts(zip: JSZip, presentationPath: string, presRels: RelationshipCollection): Promise<SlideLayout[]> {
  const layouts: SlideLayout[] = [];
  for (const masterRel of presRels.items.filter(r => r.type.includes('/slideMaster'))) {
    const masterPath = resolveRelationshipTarget(presentationPath, masterRel.target);
    const masterRelsFile = zip.file(relsPath(masterPath));
    if (!masterRelsFile) continue;
    const masterRels = new RelationshipCollection(parseXml(relsPath(masterPath), await masterRelsFile.async('string')));
    for (const layoutRel of masterRels.items.filter(r => r.type.includes('/slideLayout'))) {
      const layoutPath = resolveRelationshipTarget(masterPath, layoutRel.target);
      const file = zip.file(layoutPath); if (!file) continue;
      layouts.push(new SlideLayout(layoutPath, parseXml(layoutPath, await file.async('string')), { path: masterPath }, zip));
    }
  }
  return layouts;
}

function layoutTree(xml: XmlNode): XmlNode | undefined {
  return (((xml['p:sldLayout'] as XmlNode | undefined)?.['p:cSld'] as XmlNode | undefined)?.['p:spTree']) as XmlNode | undefined;
}
