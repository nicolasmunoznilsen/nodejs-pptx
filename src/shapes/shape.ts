import { readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';
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
  set x(value: number) { writeAttr(this.ensureXfrm()['a:off'] as XmlNode, 'x', this.assertOffset(value, 'x')); this.markDirty(); }
  set y(value: number) { writeAttr(this.ensureXfrm()['a:off'] as XmlNode, 'y', this.assertOffset(value, 'y')); this.markDirty(); }
  set width(value: number) { writeAttr(this.ensureXfrm()['a:ext'] as XmlNode, 'cx', this.assertExtent(value, 'width')); this.markDirty(); }
  set height(value: number) { writeAttr(this.ensureXfrm()['a:ext'] as XmlNode, 'cy', this.assertExtent(value, 'height')); this.markDirty(); }

  /**
   * Moves the shape to absolute EMU coordinates without changing its size.
   * Returns the shape so calls can be chained.
   */
  moveTo(x: number, y: number): this {
    this.x = x;
    this.y = y;
    return this;
  }

  /**
   * Resizes the shape to the given EMU dimensions without moving it.
   * Invalid values are rejected and the shape is left untouched.
   * Returns the shape so calls can be chained.
   */
  resize(width: number, height: number): this {
    this.width = width;
    this.height = height;
    return this;
  }

  get hasTextFrame(): boolean { return Boolean(this.node['p:txBody']); }
  get textFrame(): TextFrame {
    if (!this.hasTextFrame) throw new Error(`Shape ${this.id || this.name} does not have a text frame.`);
    return new TextFrame(this.node['p:txBody'] as XmlNode, this.markDirty);
  }
  protected get cNvPr(): XmlNode | undefined { return ((this.node['p:nvSpPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvPicPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvGraphicFramePr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvGrpSpPr'] as XmlNode | undefined)?.['p:cNvPr'] ?? (this.node['p:nvCxnSpPr'] as XmlNode | undefined)?.['p:cNvPr']) as XmlNode | undefined; }
  get xfrm(): XmlNode | undefined { return ((this.node['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] ?? (this.node['p:grpSpPr'] as XmlNode | undefined)?.['a:xfrm'] ?? (this.node['p:xfrm'] as XmlNode | undefined)) as XmlNode | undefined; }

  /** Returns the shape's transform, creating a minimal valid one when missing. */
  private ensureXfrm(): XmlNode {
    const existing = this.xfrm;
    if (existing) return existing;
    const xfrm: XmlNode = { 'a:off': { '@_x': '0', '@_y': '0' }, 'a:ext': { '@_cx': '0', '@_cy': '0' } };
    const label = this.id || this.name || this.type;
    if (this.node['p:xfrm'] !== undefined || this.type === 'graphicFrame') {
      // Graphic frames carry their transform directly as p:xfrm.
      this.node['p:xfrm'] = xfrm;
      return xfrm;
    }
    const containerKey = this.node['p:spPr'] ? 'p:spPr' : this.node['p:grpSpPr'] ? 'p:grpSpPr' : undefined;
    if (!containerKey) throw new Error(`Shape ${label} has no p:spPr, p:grpSpPr, or p:xfrm; its position and size cannot be edited.`);
    const container = this.node[containerKey] as XmlNode;
    // OOXML requires a:xfrm to be the first child of p:spPr / p:grpSpPr.
    const rebuilt: XmlNode = { 'a:xfrm': xfrm };
    for (const [key, value] of Object.entries(container)) if (key !== 'a:xfrm') rebuilt[key] = value;
    this.node[containerKey] = rebuilt;
    return xfrm;
  }

  private assertOffset(value: number, name: string): string {
    this.assertNumber(value, name);
    if (value < 0) throw new RangeError(`Shape ${this.id || this.name}: ${name} must be >= 0 EMU, got ${value}.`);
    return String(Math.round(value));
  }

  private assertExtent(value: number, name: string): string {
    this.assertNumber(value, name);
    if (value <= 0) throw new RangeError(`Shape ${this.id || this.name}: ${name} must be a positive number of EMU, got ${value}.`);
    return String(Math.round(value));
  }

  private assertNumber(value: number, name: string): void {
    if (typeof value !== 'number' || !Number.isFinite(value)) throw new TypeError(`Shape ${this.id || this.name}: ${name} must be a finite number, got ${String(value)}.`);
  }
}
