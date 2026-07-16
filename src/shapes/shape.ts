import { readAttr, type XmlNode } from '../ooxml/xml.js';
import { TextFrame } from '../text/text-frame.js';

export type ShapeType = 'textBox' | 'autoShape' | 'picture' | 'graphicFrame' | 'group' | 'connector' | 'unknown';

export class Shape {
  constructor(readonly node: XmlNode, readonly type: ShapeType, private readonly markDirty: () => void) {}
  get id(): string { return readAttr(this.cNvPr, 'id') ?? ''; }
  get name(): string { return readAttr(this.cNvPr, 'name') ?? ''; }
  get x(): number { return Number(readAttr(this.xfrm?.['a:off'], 'x') ?? 0); }
  get y(): number { return Number(readAttr(this.xfrm?.['a:off'], 'y') ?? 0); }
  get width(): number { return Number(readAttr(this.xfrm?.['a:ext'], 'cx') ?? 0); }
  get height(): number { return Number(readAttr(this.xfrm?.['a:ext'], 'cy') ?? 0); }
  get hasTextFrame(): boolean { return Boolean(this.node['p:txBody']); }
  get textFrame(): TextFrame {
    if (!this.hasTextFrame) throw new Error(`Shape ${this.id || this.name} does not have a text frame.`);
    return new TextFrame(this.node['p:txBody'] as XmlNode, this.markDirty);
  }
  protected get cNvPr(): XmlNode | undefined { return ((this.node['p:nvSpPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvPicPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvGraphicFramePr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvGrpSpPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvCxnSpPr'] as XmlNode | undefined)?.['p:cNvPr']) as XmlNode | undefined; }
  private get xfrm(): XmlNode | undefined { return ((this.node['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] ?? (this.node['p:xfrm'] as XmlNode | undefined)) as XmlNode | undefined; }
}
