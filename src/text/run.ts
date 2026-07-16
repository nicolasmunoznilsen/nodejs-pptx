import { readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';

export interface RunFont { name?: string; size?: number; bold?: boolean; italic?: boolean; color?: string }

export class Run {
  readonly font: RunFont;
  constructor(private readonly node: XmlNode, private readonly markDirty: () => void) {
    this.font = createFontProxy(this.rPr, markDirty);
  }
  get text(): string { return typeof this.node['a:t'] === 'string' || typeof this.node['a:t'] === 'number' ? String(this.node['a:t']) : ''; }
  set text(value: string) { if (this.text !== value) { this.node['a:t'] = value; this.markDirty(); } }
  get rawNode(): XmlNode { return this.node; }
  private get rPr(): XmlNode { if (!this.node['a:rPr'] || typeof this.node['a:rPr'] !== 'object') this.node['a:rPr'] = {}; return this.node['a:rPr'] as XmlNode; }
}

function createFontProxy(rPr: XmlNode, markDirty: () => void): RunFont {
  return {
    get name() { return readAttr(rPr['a:latin'], 'typeface'); },
    set name(value) { ensureChild(rPr, 'a:latin'); writeAttr(rPr['a:latin'] as XmlNode, 'typeface', value); markDirty(); },
    get size() { const raw = readAttr(rPr, 'sz'); return raw ? Number(raw) / 100 : undefined; },
    set size(value) { writeAttr(rPr, 'sz', value === undefined ? undefined : String(value * 100)); markDirty(); },
    get bold() { return readAttr(rPr, 'b') === '1'; },
    set bold(value) { writeAttr(rPr, 'b', value ? '1' : undefined); markDirty(); },
    get italic() { return readAttr(rPr, 'i') === '1'; },
    set italic(value) { writeAttr(rPr, 'i', value ? '1' : undefined); markDirty(); },
    get color() { return readAttr(((rPr['a:solidFill'] as XmlNode | undefined)?.['a:srgbClr']), 'val'); },
    set color(value) {
      if (!value) { delete rPr['a:solidFill']; markDirty(); return; }
      rPr['a:solidFill'] = { 'a:srgbClr': { '@_val': value.replace('#', '').toUpperCase() } };
      markDirty();
    },
  } as RunFont;
}

function ensureChild(node: XmlNode, key: string): void { if (!node[key] || typeof node[key] !== 'object') node[key] = {}; }
