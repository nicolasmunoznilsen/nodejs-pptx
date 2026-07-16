import { access, readFile, writeFile } from 'node:fs/promises';
import { basename } from 'node:path';
import JSZip from 'jszip';
import { MAIN_REL, REL_NS, SLIDE_LAYOUT_REL, SLIDE_MASTER_REL, SLIDE_REL, THEME_REL } from '../ooxml/constants.js';
import { allocateSlidePath } from '../ooxml/ids.js';
import { relsPath, resolveRelationshipTarget } from '../ooxml/paths.js';
import { arr, parseXml, readAttr, xmlToString, type XmlNode } from '../ooxml/xml.js';
import { ContentTypes } from '../package/content-types.js';
import { RelationshipCollection } from '../package/relationships.js';
import { blankSlideXml, Slide } from '../slides/slide.js';
import { SlideCollection } from './slide-collection.js';

export interface PresentationOptions { title?: string; author?: string; width?: number; height?: number }

export class Presentation {
  readonly slides = new SlideCollection(this);
  private zip: JSZip;
  private presentationPath = 'ppt/presentation.xml';
  private presentationXml: XmlNode;
  private presentationDirty = false;
  private presentationRelsPath = 'ppt/_rels/presentation.xml.rels';
  private presentationRels: RelationshipCollection;
  private contentTypes: ContentTypes;
  readonly slideList: Slide[] = [];

  private constructor(zip: JSZip, presentationXml: XmlNode, presentationRels: RelationshipCollection, contentTypes: ContentTypes) {
    this.zip = zip;
    this.presentationXml = presentationXml;
    this.presentationRels = presentationRels;
    this.contentTypes = contentTypes;
  }

  static create(options: PresentationOptions = {}): Presentation {
    const zip = new JSZip();
    const contentTypes = ContentTypes.create();
    contentTypes.addOverride('ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml');
    contentTypes.addOverride('ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml');
    contentTypes.addOverride('ppt/slideLayouts/slideLayout1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml');
    contentTypes.addOverride('ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
    contentTypes.addOverride('docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
    contentTypes.save(zip, true);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${MAIN_REL}" Target="ppt/presentation.xml"/></Relationships>`);
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXml(options.title ?? 'Untitled presentation')}</dc:title><dc:creator>${escapeXml(options.author ?? 'nodejs-pptx')}</dc:creator></cp:coreProperties>`);
    const presXml = parseXml('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst></p:sldIdLst><p:sldSz cx="${inches(options.width ?? 10)}" cy="${inches(options.height ?? 5.625)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
    const presRoot = presXml['p:presentation'] as XmlNode;
    if (typeof presRoot['p:sldIdLst'] === 'string') presRoot['p:sldIdLst'] = {};
    const presRels = RelationshipCollection.empty();
    presRels.add(SLIDE_MASTER_REL, 'slideMasters/slideMaster1.xml');
    zip.file('ppt/theme/theme1.xml', themeXml());
    zip.file('ppt/slideLayouts/slideLayout1.xml', slideLayoutXml());
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${SLIDE_MASTER_REL}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
    zip.file('ppt/slideMasters/slideMaster1.xml', slideMasterXml());
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${SLIDE_LAYOUT_REL}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${THEME_REL}" Target="../theme/theme1.xml"/></Relationships>`);
    const p = new Presentation(zip, presXml, presRels, contentTypes);
    p.presentationDirty = true;
    p.presentationRels.dirty = true;
    p.persistPresentationIfDirty();
    return p;
  }

  static async open(input: string | Buffer): Promise<Presentation> {
    let data: Buffer;
    if (typeof input === 'string') { try { await access(input); } catch { throw new Error(`PPTX file does not exist: ${input}`); } data = await readFile(input); } else data = input;
    let zip: JSZip;
    try { zip = await JSZip.loadAsync(data); } catch (e) { throw new Error(`Invalid PPTX zip package: ${(e as Error).message}`); }
    const rootRelsFile = zip.file('_rels/.rels');
    if (!rootRelsFile) throw new Error('Invalid PPTX: missing _rels/.rels.');
    const rootRels = new RelationshipCollection(parseXml('_rels/.rels', await rootRelsFile.async('string')));
    const mainRel = rootRels.findByType(MAIN_REL) ?? rootRels.items.find(r => r.target.endsWith('presentation.xml'));
    if (!mainRel) throw new Error('Invalid PPTX: missing officeDocument relationship.');
    const presPath = resolveRelationshipTarget('_rels/.rels', mainRel.target);
    const presFile = zip.file(presPath); if (!presFile) throw new Error(`Invalid PPTX: missing ${presPath}.`);
    const presXml = parseXml(presPath, await presFile.async('string'));
    const presRelsPath = relsPath(presPath); const presRelsFile = zip.file(presRelsPath); if (!presRelsFile) throw new Error(`Invalid PPTX: missing ${presRelsPath}.`);
    const presRels = new RelationshipCollection(parseXml(presRelsPath, await presRelsFile.async('string')));
    const contentTypes = await ContentTypes.load(zip);
    const p = new Presentation(zip, presXml, presRels, contentTypes); p.presentationPath = presPath; p.presentationRelsPath = presRelsPath;
    await p.loadSlides();
    return p;
  }

  async save(filePath: string): Promise<void> { await writeFile(filePath, await this.toBuffer()); }
  async toBuffer(): Promise<Buffer> {
    this.persistPresentationIfDirty();
    this.contentTypes.save(this.zip);
    this.slideList.forEach(s => s.persistIfDirty());
    return await this.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }

  /** @internal */
  addSlideInternal(): Slide {
    const slidePath = allocateSlidePath(this.zip);
    const slideXml = parseXml(slidePath, blankSlideXml());
    const slideRels = RelationshipCollection.empty();
    slideRels.add(SLIDE_LAYOUT_REL, '../slideLayouts/slideLayout1.xml');
    const slide = new Slide(this.zip, slidePath, slideXml, slideRels.xml);
    this.slideList.push(slide);
    const rel = this.presentationRels.add(SLIDE_REL, `slides/${basename(slidePath)}`);
    const sldIdLst = (this.presentationXml['p:presentation'] as XmlNode)['p:sldIdLst'] as XmlNode;
    const ids = arr(sldIdLst['p:sldId'] as XmlNode | XmlNode[] | undefined);
    const maxSldId = ids.reduce((max, slideId) => Math.max(max, parseInt(readAttr(slideId, 'id') || '0')), 255);
    ids.push({ '@_id': String(maxSldId + 1), '@_r:id': rel.id });
    sldIdLst['p:sldId'] = ids;
    this.contentTypes.addOverride(slidePath, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    slide.persist();
    this.zip.file(relsPath(slidePath), xmlToString(slideRels.xml));
    this.presentationDirty = true;
    return slide;
  }

  private persistPresentationIfDirty(): void {
    if (this.presentationDirty) this.zip.file(this.presentationPath, xmlToString(this.presentationXml));
    if (this.presentationRels.dirty) this.zip.file(this.presentationRelsPath, xmlToString(this.presentationRels.xml));
  }

  private async loadSlides(): Promise<void> {
    const pres = this.presentationXml['p:presentation'] as XmlNode | undefined; let list = pres?.['p:sldIdLst'] as XmlNode | string | undefined;
    if (!pres) throw new Error('Invalid PPTX: missing p:presentation or p:sldIdLst.');
    if (typeof list === 'string') { list = {}; pres['p:sldIdLst'] = list; }
    if (!list) throw new Error('Invalid PPTX: missing p:presentation or p:sldIdLst.');
    for (const sldId of arr(list['p:sldId'] as XmlNode | XmlNode[] | undefined)) {
      const rid = readAttr(sldId, 'r:id'); const rel = rid ? this.presentationRels.getById(rid) : undefined;
      if (!rid || !rel || rel.type !== SLIDE_REL) throw new Error(`Invalid PPTX: missing slide relationship for ${rid ?? 'unknown id'}.`);
      const slidePath = resolveRelationshipTarget(this.presentationPath, rel.target);
      const file = this.zip.file(slidePath); if (!file) throw new Error(`Invalid PPTX: missing slide part ${slidePath}.`);
      const relsFile = this.zip.file(relsPath(slidePath));
      this.slideList.push(Slide.parse(this.zip, slidePath, await file.async('string'), relsFile ? await relsFile.async('string') : undefined));
    }
  }
}

function escapeXml(value: string): string { return value.replace(/[<>&"']/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' })[c]!); }
export function inches(value: number): number { return Math.round(value * 914400); }
function slideLayoutXml(): string { return `<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`; }
function slideMasterXml(): string { return `<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`; }
function themeXml(): string { return `<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="nodejs-pptx"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`; }
