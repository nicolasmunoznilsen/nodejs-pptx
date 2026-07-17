import type JSZip from 'jszip';
import { nextShapeId } from '../ooxml/ids.js';
import { relsPath } from '../ooxml/paths.js';
import { arr, parseXml, readAttr, walk, xmlToString, type XmlNode } from '../ooxml/xml.js';
import { RelationshipCollection } from '../package/relationships.js';
import { ShapeCollection } from '../shapes/shape-collection.js';
import { textBoxXml, type TextBoxInfo, type TextBoxOptions } from '../shapes/text-box.js';
import type { Shape } from '../shapes/shape.js';

export class Slide {
  dirty = false;
  relsDirty = false;
  private readonly relationships: RelationshipCollection;
  readonly shapes: ShapeCollection;

  constructor(private readonly zip: JSZip, readonly path: string, private xml: XmlNode, relsXml?: XmlNode) {
    this.relationships = new RelationshipCollection(relsXml);
    this.shapes = new ShapeCollection(this.spTree, () => this.markDirty());
  }

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

  persist(): void { this.markDirty(); this.zip.file(this.path, xmlToString(this.xml)); }
  persistIfDirty(): void {
    if (this.dirty) this.zip.file(this.path, xmlToString(this.xml));
    if (this.relationships.dirty || this.relsDirty) this.zip.file(relsPath(this.path), xmlToString(this.relationships.xml));
  }

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
