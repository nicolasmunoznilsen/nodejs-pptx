import { arr, type XmlNode } from '../ooxml/xml.js';
import { Shape, type ShapeType } from './shape.js';

export class ShapeCollection {
  constructor(private readonly spTree: XmlNode, private readonly markDirty: () => void) {}
  get length(): number { return this.shapes.length; }
  get(index: number): Shape {
    if (!Number.isInteger(index) || index < 0 || index >= this.length) throw new RangeError(`Shape index ${index} is out of range (0-${this.length - 1}).`);
    return this.shapes[index];
  }
  findByName(name: string): Shape | undefined { return this.shapes.find(shape => shape.name === name); }
  findById(id: string): Shape | undefined { return this.shapes.find(shape => shape.id === id); }
  [Symbol.iterator](): IterableIterator<Shape> { return this.shapes[Symbol.iterator](); }
  get all(): Shape[] { return this.shapes; }

  private get shapes(): Shape[] {
    const out: Shape[] = [];
    for (const [key, type] of [['p:sp', 'autoShape'], ['p:pic', 'picture'], ['p:graphicFrame', 'graphicFrame'], ['p:grpSp', 'group'], ['p:cxnSp', 'connector']] as Array<[string, ShapeType]>) {
      for (const node of arr(this.spTree[key] as XmlNode | XmlNode[] | undefined)) {
        const finalType = key === 'p:sp' && (node['p:nvSpPr'] as XmlNode | undefined)?.['p:cNvSpPr'] && ((node['p:nvSpPr'] as XmlNode)['p:cNvSpPr'] as XmlNode)['@_txBox'] === '1' ? 'textBox' : type;
        out.push(new Shape(node, finalType, this.markDirty));
      }
    }
    return out;
  }
}
