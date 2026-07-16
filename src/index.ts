import { access, readFile, writeFile } from 'node:fs/promises';
import { basename } from 'node:path';
import JSZip from 'jszip';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';

export const EMU_PER_INCH = 914400;
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const MAIN_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

type XmlNode = Record<string, unknown>;
type Search = string | RegExp;
export interface TextBoxOptions { x?: number; y?: number; width?: number; height?: number; fontSize?: number; color?: string; name?: string }
export interface TextBoxInfo { text: string; name: string; id: string; index: number; x: number; y: number; width: number; height: number }
export interface PresentationOptions { title?: string; author?: string; width?: number; height?: number }

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', preserveOrder: false, trimValues: false });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', suppressEmptyNode: true, format: false });

export function inches(value: number): number { return Math.round(value * EMU_PER_INCH); }
function escapeXml(value: string): string { return value.replace(/[<>&"']/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' })[c]!); }
function textOf(v: unknown): string { return typeof v === 'string' || typeof v === 'number' ? String(v) : ''; }
function arr<T>(value: T | T[] | undefined): T[] { return value === undefined ? [] : Array.isArray(value) ? value : [value]; }
function localTarget(baseDir: string, target: string): string { return target.startsWith('/') ? target.slice(1) : `${baseDir}/${target}`.replaceAll('//', '/'); }
function relsPath(part: string): string { const i = part.lastIndexOf('/'); return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`; }
function xmlToString(xml: XmlNode): string { return builder.build(xml); }
function parseXml(name: string, xml: string): XmlNode { try { return parser.parse(xml) as XmlNode; } catch (e) { throw new Error(`Invalid XML in ${name}: ${(e as Error).message}`); } }
function readAttr(node: unknown, attr: string): string | undefined { return typeof node === 'object' && node !== null ? (node as XmlNode)[`@_${attr}`] as string | undefined : undefined; }
function walk(value: unknown, visit: (node: XmlNode) => void): void { if (!value || typeof value !== 'object') return; if (Array.isArray(value)) { value.forEach(v => walk(v, visit)); return; } const node = value as XmlNode; visit(node); Object.values(node).forEach(v => walk(v, visit)); }
function relArray(relsXml: XmlNode): XmlNode[] { return arr((relsXml.Relationships as XmlNode | undefined)?.Relationship as XmlNode | XmlNode[] | undefined); }

export class SlideCollection {
  constructor(private readonly presentation: Presentation) {}
  get length(): number { return this.presentation.slideList.length; }
  get(index: number): Slide {
    if (!Number.isInteger(index) || index < 0 || index >= this.length) throw new RangeError(`Slide index ${index} is out of range (0-${this.length - 1}).`);
    return this.presentation.slideList[index];
  }
  add(): Slide { return this.presentation.addSlideInternal(); }
  [Symbol.iterator](): IterableIterator<Slide> { return this.presentation.slideList[Symbol.iterator](); }
}

export class Presentation {
  readonly slides = new SlideCollection(this);
  private zip: JSZip;
  private presentationPath = 'ppt/presentation.xml';
  private presentationXml: XmlNode;
  private presentationRelsPath = 'ppt/_rels/presentation.xml.rels';
  private presentationRelsXml: XmlNode;
  readonly slideList: Slide[] = [];

  private constructor(zip: JSZip, presentationXml: XmlNode, presentationRelsXml: XmlNode) {
    this.zip = zip;
    this.presentationXml = presentationXml;
    this.presentationRelsXml = presentationRelsXml;
  }

  static create(options: PresentationOptions = {}): Presentation {
    const zip = new JSZip();
    const title = escapeXml(options.title ?? 'Untitled presentation');
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${MAIN_REL}" Target="ppt/presentation.xml"/></Relationships>`);
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${title}</dc:title><dc:creator>${escapeXml(options.author ?? 'nodejs-pptx')}</dc:creator></cp:coreProperties>`);
    const presXml = parseXml('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst></p:sldIdLst><p:sldSz cx="${inches(options.width ?? 10)}" cy="${inches(options.height ?? 5.625)}"/></p:presentation>`);
    const relsXml = parseXml('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL_NS}"></Relationships>`);
    const p = new Presentation(zip, presXml, relsXml);
    p.persistPresentation();
    return p;
  }

  static async open(input: string | Buffer): Promise<Presentation> {
    let data: Buffer;
    if (typeof input === 'string') { try { await access(input); } catch { throw new Error(`PPTX file does not exist: ${input}`); } data = await readFile(input); } else data = input;
    let zip: JSZip;
    try { zip = await JSZip.loadAsync(data); } catch (e) { throw new Error(`Invalid PPTX zip package: ${(e as Error).message}`); }
    const rootRelsFile = zip.file('_rels/.rels');
    if (!rootRelsFile) throw new Error('Invalid PPTX: missing _rels/.rels.');
    const rootRels = parseXml('_rels/.rels', await rootRelsFile.async('string'));
    const mainRel = relArray(rootRels).find(r => readAttr(r, 'Type') === MAIN_REL) ?? relArray(rootRels).find(r => readAttr(r, 'Target')?.endsWith('presentation.xml'));
    if (!mainRel) throw new Error('Invalid PPTX: missing officeDocument relationship.');
    const presPath = localTarget('', readAttr(mainRel, 'Target') ?? 'ppt/presentation.xml');
    const presFile = zip.file(presPath); if (!presFile) throw new Error(`Invalid PPTX: missing ${presPath}.`);
    const presXml = parseXml(presPath, await presFile.async('string'));
    const presRelsPath = relsPath(presPath); const presRelsFile = zip.file(presRelsPath); if (!presRelsFile) throw new Error(`Invalid PPTX: missing ${presRelsPath}.`);
    const presRelsXml = parseXml(presRelsPath, await presRelsFile.async('string'));
    const p = new Presentation(zip, presXml, presRelsXml); p.presentationPath = presPath; p.presentationRelsPath = presRelsPath;
    await p.loadSlides();
    return p;
  }

  async save(filePath: string): Promise<void> { await writeFile(filePath, await this.toBuffer()); }
  async toBuffer(): Promise<Buffer> { this.persistPresentation(); this.slideList.forEach(s => s.persist()); await ensureSlideContentTypes(this.zip, this.slideList.map(slide => slide.path)); return await this.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }); }

  /** @internal */
  addSlideInternal(): Slide {
    const number = this.slideList.length + 1;
    const slidePath = `ppt/slides/slide${number}.xml`;
    const slideXml = parseXml(slidePath, blankSlideXml());
    const slide = new Slide(this.zip, slidePath, slideXml);
    this.slideList.push(slide);
    const rels = this.relationships();
    const relId = `rId${rels.length + 1}`;
    rels.push({ '@_Id': relId, '@_Type': SLIDE_REL, '@_Target': `slides/${basename(slidePath)}` });
    const sldIdLst = (this.presentationXml['p:presentation'] as XmlNode)['p:sldIdLst'] as XmlNode;
    const ids = arr(sldIdLst['p:sldId'] as XmlNode | XmlNode[] | undefined);
    ids.push({ '@_id': String(256 + number), '@_r:id': relId });
    sldIdLst['p:sldId'] = ids;
    this.zip.file(slidePath, xmlToString(slideXml));
    this.zip.file(relsPath(slidePath), `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL_NS}"></Relationships>`);
    return slide;
  }

  private relationships(): XmlNode[] { const root = this.presentationRelsXml.Relationships as XmlNode; const rels = relArray(this.presentationRelsXml); root.Relationship = rels; return rels; }
  private persistPresentation(): void { this.zip.file(this.presentationPath, xmlToString(this.presentationXml)); this.zip.file(this.presentationRelsPath, xmlToString(this.presentationRelsXml)); }
  private async loadSlides(): Promise<void> {
    const pres = this.presentationXml['p:presentation'] as XmlNode | undefined; const list = pres?.['p:sldIdLst'] as XmlNode | undefined;
    if (!pres || !list) throw new Error('Invalid PPTX: missing p:presentation or p:sldIdLst.');
    const rels = relArray(this.presentationRelsXml);
    for (const sldId of arr(list['p:sldId'] as XmlNode | XmlNode[] | undefined)) {
      const rid = readAttr(sldId, 'r:id'); const rel = rels.find(r => readAttr(r, 'Id') === rid && readAttr(r, 'Type') === SLIDE_REL);
      if (!rid || !rel) throw new Error(`Invalid PPTX: missing slide relationship for ${rid ?? 'unknown id'}.`);
      const slidePath = localTarget(this.presentationPath.slice(0, this.presentationPath.lastIndexOf('/')), readAttr(rel, 'Target') ?? '');
      const file = this.zip.file(slidePath); if (!file) throw new Error(`Invalid PPTX: missing slide part ${slidePath}.`);
      this.slideList.push(new Slide(this.zip, slidePath, parseXml(slidePath, await file.async('string'))));
    }
  }
}

export class Slide {
  constructor(private readonly zip: JSZip, readonly path: string, private xml: XmlNode) {}
  get text(): string { return this.getTextBoxes().map(t => t.text).filter(Boolean).join('\n'); }
  getTextBoxes(): TextBoxInfo[] {
    const boxes: TextBoxInfo[] = [];
    walk(this.xml, node => {
      const sp = node['p:sp'] as XmlNode | undefined; if (!sp?.['p:txBody']) return;
      const nv = sp['p:nvSpPr'] as XmlNode | undefined; const cNvPr = nv?.['p:cNvPr'] as XmlNode | undefined;
      const xfrm = (sp['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] as XmlNode | undefined;
      boxes.push({ text: collectText(sp), name: readAttr(cNvPr, 'name') ?? '', id: readAttr(cNvPr, 'id') ?? String(boxes.length), index: boxes.length, x: Number(readAttr(xfrm?.['a:off'], 'x') ?? 0), y: Number(readAttr(xfrm?.['a:off'], 'y') ?? 0), width: Number(readAttr(xfrm?.['a:ext'], 'cx') ?? 0), height: Number(readAttr(xfrm?.['a:ext'], 'cy') ?? 0) });
    });
    return boxes;
  }
  replaceText(search: Search, replacement: string): number {
    let count = 0;
    walk(this.xml, node => { if (Object.prototype.hasOwnProperty.call(node, 'a:t')) { const before = textOf(node['a:t']); const after = typeof search === 'string' ? before.split(search).join(replacement) : before.replace(search, replacement); if (after !== before) { count += 1; node['a:t'] = after; } } });
    if (count) this.persist();
    return count;
  }
  addTextBox(text: string, options: TextBoxOptions = {}): TextBoxInfo {
    const tree = slideTree(this.xml); const shapes = arr(tree['p:sp'] as XmlNode | XmlNode[] | undefined); const id = String(shapes.length + 2);
    const shape = textBoxXml(text, id, options); shapes.push(shape); tree['p:sp'] = shapes; this.persist(); return this.getTextBoxes().at(-1)!;
  }
  persist(): void { this.zip.file(this.path, xmlToString(this.xml)); }
}

function collectText(node: unknown): string { const parts: string[] = []; walk(node, n => { if (Object.prototype.hasOwnProperty.call(n, 'a:t')) parts.push(textOf(n['a:t'])); }); return parts.join(''); }
function slideTree(xml: XmlNode): XmlNode { const tree = (((xml['p:sld'] as XmlNode)['p:cSld'] as XmlNode)['p:spTree']) as XmlNode | undefined; if (!tree) throw new Error('Invalid slide XML: missing p:spTree.'); return tree; }
function textBoxXml(text: string, id: string, o: TextBoxOptions): XmlNode { return { 'p:nvSpPr': { 'p:cNvPr': { '@_id': id, '@_name': o.name ?? `TextBox ${id}` }, 'p:cNvSpPr': { '@_txBox': '1' }, 'p:nvPr': {} }, 'p:spPr': { 'a:xfrm': { 'a:off': { '@_x': String(inches(o.x ?? 0.5)), '@_y': String(inches(o.y ?? 0.5)) }, 'a:ext': { '@_cx': String(inches(o.width ?? 4)), '@_cy': String(inches(o.height ?? 1)) } }, 'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} }, 'a:noFill': {}, 'a:ln': { 'a:noFill': {} } }, 'p:txBody': { 'a:bodyPr': { '@_wrap': 'square' }, 'a:lstStyle': {}, 'a:p': { 'a:r': { 'a:rPr': { '@_lang': 'en-US', '@_sz': String((o.fontSize ?? 18) * 100), 'a:solidFill': { 'a:srgbClr': { '@_val': (o.color ?? '000000').replace('#', '').toUpperCase() } } }, 'a:t': text } } } }; }
function blankSlideXml(): string { return `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`; }
async function ensureSlideContentTypes(zip: JSZip, slidePaths: string[]): Promise<void> { const f = zip.file('[Content_Types].xml'); if (!f) throw new Error('Invalid PPTX: missing [Content_Types].xml.'); let xml = await f.async('string'); for (const slidePath of slidePaths) { if (!xml.includes(`/${slidePath}`)) xml = xml.replace('</Types>', `<Override PartName="/${slidePath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`); } zip.file('[Content_Types].xml', xml); }
