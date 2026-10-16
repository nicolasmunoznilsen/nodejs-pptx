import { readAttr, type XmlNode } from '../ooxml/xml.js';
import { EMU_PER_INCH } from '../ooxml/constants.js';
import { TextFrame } from '../text/text-frame.js';
import { contentTypeForExtension, type ImageFormat } from './picture.js';

function emuToInches(value: string | undefined): number {
  const emu = Number(value ?? 0);
  return Number.isFinite(emu) ? emu / EMU_PER_INCH : 0;
}

export type ShapeType = 'textBox' | 'autoShape' | 'picture' | 'graphicFrame' | 'group' | 'connector' | 'unknown';

export class Shape {
  constructor(readonly node: XmlNode, readonly type: ShapeType, private readonly markDirty: () => void, private readonly resolveTarget?: (relationshipId: string) => string | undefined) {}
  get id(): string { return readAttr(this.cNvPr, 'id') ?? ''; }
  get name(): string { return readAttr(this.cNvPr, 'name') ?? ''; }
  /** Horizontal position of the shape, in inches (EMU in the package XML). */
  get x(): number { return emuToInches(readAttr(this.xfrm?.['a:off'], 'x')); }
  /** Vertical position of the shape, in inches (EMU in the package XML). */
  get y(): number { return emuToInches(readAttr(this.xfrm?.['a:off'], 'y')); }
  /** Width of the shape, in inches (EMU in the package XML). */
  get width(): number { return emuToInches(readAttr(this.xfrm?.['a:ext'], 'cx')); }
  /** Height of the shape, in inches (EMU in the package XML). */
  get height(): number { return emuToInches(readAttr(this.xfrm?.['a:ext'], 'cy')); }
  get hasTextFrame(): boolean { return Boolean(this.node['p:txBody']); }
  get hasImage(): boolean { return Boolean(this.blip); }
  /** Relationship id of the embedded image, when this shape is a picture. */
  get imageRelationshipId(): string | undefined { return readAttr(this.blip, 'embed'); }
  /** Package part name of the embedded image, when it can be resolved. */
  get imageFileName(): string | undefined {
    const rid = this.imageRelationshipId;
    return rid && this.resolveTarget ? this.resolveTarget(rid) : undefined;
  }
  get imageContentType(): string | undefined {
    const name = this.imageFileName;
    if (!name) return undefined;
    return contentTypeForExtension(name.slice(name.lastIndexOf('.') + 1));
  }
  get imageFormat(): ImageFormat | undefined {
    const type = this.imageContentType;
    return type === 'image/png' ? 'png' : type === 'image/jpeg' ? 'jpeg' : undefined;
  }
  get textFrame(): TextFrame {
    if (!this.hasTextFrame) throw new Error(`Shape ${this.id || this.name} does not have a text frame.`);
    return new TextFrame(this.node['p:txBody'] as XmlNode, this.markDirty);
  }
  protected get cNvPr(): XmlNode | undefined { return ((this.node['p:nvSpPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvPicPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvGraphicFramePr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvGrpSpPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvCxnSpPr'] as XmlNode | undefined)?.['p:cNvPr']) as XmlNode | undefined; }
  private get xfrm(): XmlNode | undefined { return ((this.node['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] ?? (this.node['p:xfrm'] as XmlNode | undefined)) as XmlNode | undefined; }
  private get blip(): XmlNode | undefined {
    if (this.type !== 'picture') return undefined;
    return ((this.node['p:blipFill'] as XmlNode | undefined)?.['a:blip']) as XmlNode | undefined;
  }
}
