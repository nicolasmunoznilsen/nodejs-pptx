import { TextFrame } from '../text/text-frame.js';
import { arr, readAttr, type XmlNode } from '../ooxml/xml.js';

export type PlaceholderType = 'title'|'centerTitle'|'subtitle'|'body'|'object'|'picture'|'date'|'footer'|'slideNumber'|'unknown';

export class Placeholder {
  constructor(private readonly shape: XmlNode, private readonly markDirty: () => void) {}
  get id(): string { return readAttr(this.cnv, 'id') ?? ''; }
  get name(): string { return readAttr(this.cnv, 'name') ?? ''; }
  get type(): PlaceholderType { const t = readAttr(this.ph, 'type') ?? (readAttr(this.ph, 'idx') ? 'body' : 'object'); return mapPlaceholderType(t); }
  get idx(): string | undefined { return readAttr(this.ph, 'idx'); }
  get x(): number { return num(this.off, 'x'); }
  get y(): number { return num(this.off, 'y'); }
  get width(): number { return num(this.ext, 'cx'); }
  get height(): number { return num(this.ext, 'cy'); }
  get hasTextFrame(): boolean { return !!this.shape['p:txBody']; }
  get textFrame(): TextFrame { if (!this.shape['p:txBody']) this.shape['p:txBody'] = { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': { 'a:r': { 'a:t': '' } } }; return new TextFrame(this.shape['p:txBody'] as XmlNode, this.markDirty); }
  private get nvSpPr(): XmlNode { return (this.shape['p:nvSpPr'] as XmlNode) ?? {}; }
  private get cnv(): XmlNode { return (this.nvSpPr['p:cNvPr'] as XmlNode) ?? {}; }
  private get ph(): XmlNode { return (((this.nvSpPr['p:nvPr'] as XmlNode) ?? {})['p:ph'] as XmlNode) ?? {}; }
  private get xfrm(): XmlNode { return (((this.shape['p:spPr'] as XmlNode) ?? {})['a:xfrm'] as XmlNode) ?? {}; }
  private get off(): XmlNode { return (this.xfrm['a:off'] as XmlNode) ?? {}; }
  private get ext(): XmlNode { return (this.xfrm['a:ext'] as XmlNode) ?? {}; }
}

export class PlaceholderCollection {
  constructor(private readonly items: Placeholder[]) {}
  get length(): number { return this.items.length; }
  get(index: number): Placeholder { if (index < 0 || index >= this.items.length) throw new RangeError(`Placeholder index ${index} is out of range.`); return this.items[index]; }
  findByType(type: PlaceholderType): Placeholder | undefined { return this.items.find(p => p.type === type); }
  findByName(name: string): Placeholder | undefined { return this.items.find(p => p.name === name); }
  toArray(): Placeholder[] { return [...this.items]; }
  [Symbol.iterator](): IterableIterator<Placeholder> { return this.items[Symbol.iterator](); }
}

export function placeholdersFromTree(spTree: XmlNode | undefined, markDirty: () => void): Placeholder[] {
  return arr(spTree?.['p:sp'] as XmlNode | XmlNode[] | undefined).filter(s => !!(((s['p:nvSpPr'] as XmlNode|undefined)?.['p:nvPr'] as XmlNode|undefined)?.['p:ph'])).map(s => new Placeholder(s, markDirty));
}
function num(node: XmlNode, attr: string): number { return parseInt(readAttr(node, attr) ?? '0', 10) || 0; }

/**
 * `p:ph` `type` attribute value -> public placeholder type. Both the OOXML spellings
 * (`ctrTitle`, `subTitle`, `sldNum`, ...) and the public spellings of `PlaceholderType` are
 * accepted, matching what other tools write into real decks.
 */
const OOXML_TO_PLACEHOLDER_TYPE: Readonly<Record<string, PlaceholderType>> = {
  title: 'title',
  ctrTitle: 'centerTitle',
  centerTitle: 'centerTitle',
  subTitle: 'subtitle',
  subtitle: 'subtitle',
  body: 'body',
  object: 'object',
  pic: 'picture',
  picture: 'picture',
  dt: 'date',
  date: 'date',
  ftr: 'footer',
  footer: 'footer',
  sldNum: 'slideNumber',
  slideNumber: 'slideNumber',
};

/**
 * Maps a `p:ph` `type` attribute value to the public placeholder type.
 *
 * The lookup is guarded with `Object.hasOwn`, so names inherited from `Object.prototype`
 * (`constructor`, `toString`, `valueOf`, `__proto__`, ...) never resolve through the prototype
 * chain: a document containing `type="constructor"` maps to `'unknown'` instead of leaking an
 * inherited value. Unknown and malformed names map to `'unknown'` too.
 */
export function mapPlaceholderType(rawType: string): PlaceholderType {
  if (typeof rawType !== 'string' || !Object.hasOwn(OOXML_TO_PLACEHOLDER_TYPE, rawType)) return 'unknown';
  return OOXML_TO_PLACEHOLDER_TYPE[rawType];
}
