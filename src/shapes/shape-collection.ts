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

  /**
   * Every shape in document order, descending into the shape trees nested inside `p:grpSp` groups.
   * Group children are deliberately absent from `all`/`get(index)` (they are not siblings of the
   * group), but text replacement has to reach them: otherwise a deck whose text boxes were grouped
   * in PowerPoint would silently keep the old text.
   */
  get allDescendants(): Shape[] {
    const out: Shape[] = [];
    this.collectDescendants(this.spTree, out);
    return out;
  }

  private collectDescendants(tree: XmlNode, out: Shape[]): void {
    for (const shape of this.shapesIn(tree)) {
      out.push(shape);
      const nested = shape.node['p:spTree'];
      if (nested && typeof nested === 'object') this.collectDescendants(nested as XmlNode, out);
    }
  }

  private shapesIn(tree: XmlNode): Shape[] {
    const out: Shape[] = [];
    for (const [key, type] of [['p:sp', 'autoShape'], ['p:pic', 'picture'], ['p:graphicFrame', 'graphicFrame'], ['p:grpSp', 'group'], ['p:cxnSp', 'connector']] as Array<[string, ShapeType]>) {
      for (const node of arr(tree[key] as XmlNode | XmlNode[] | undefined)) {
        const finalType = key === 'p:sp' && (node['p:nvSpPr'] as XmlNode | undefined)?.['p:cNvSpPr'] && ((node['p:nvSpPr'] as XmlNode)['p:cNvSpPr'] as XmlNode)['@_txBox'] === '1' ? 'textBox' : type;
        out.push(new Shape(node, finalType, this.markDirty));
      }
    }
    return out;
  }

  private get shapes(): Shape[] {
    return this.shapesIn(this.spTree);
  }
}
