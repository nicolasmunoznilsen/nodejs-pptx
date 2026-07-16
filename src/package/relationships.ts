import { arr, parseXml, readAttr, type XmlNode } from '../ooxml/xml.js';
import { maxNumericRelationshipId } from '../ooxml/ids.js';
import { REL_NS } from '../ooxml/constants.js';

export interface Relationship { id: string; type: string; target: string }

export class RelationshipCollection {
  dirty = false;
  readonly xml: XmlNode;

  constructor(xml?: XmlNode) {
    this.xml = xml ?? parseXml('relationships', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL_NS}"></Relationships>`);
    if (typeof this.xml.Relationships === 'string') this.xml.Relationships = {};
    if (!this.xml.Relationships) this.xml.Relationships = { '@_xmlns': REL_NS };
  }

  static empty(): RelationshipCollection { return new RelationshipCollection(); }

  get items(): Relationship[] { return this.nodes.map(nodeToRelationship); }
  getById(id: string): Relationship | undefined { return this.items.find(r => r.id === id); }
  findByType(type: string): Relationship | undefined { return this.items.find(r => r.type === type); }
  findByTarget(target: string): Relationship | undefined { return this.items.find(r => r.target === target); }
  allocateId(): string { return `rId${maxNumericRelationshipId(this.items) + 1}`; }

  add(type: string, target: string): Relationship {
    const existing = this.items.find(r => r.type === type && r.target === target);
    if (existing) return existing;
    const rel = { id: this.allocateId(), type, target };
    this.nodes = [...this.nodes, relationshipToNode(rel)];
    this.dirty = true;
    return rel;
  }

  remove(id: string): void {
    const before = this.nodes;
    const after = before.filter(node => readAttr(node, 'Id') !== id);
    if (after.length !== before.length) { this.nodes = after; this.dirty = true; }
  }

  private get root(): XmlNode { return this.xml.Relationships as XmlNode; }
  private get nodes(): XmlNode[] { return arr(this.root.Relationship as XmlNode | XmlNode[] | undefined); }
  private set nodes(nodes: XmlNode[]) { this.root.Relationship = nodes; }
}

function nodeToRelationship(node: XmlNode): Relationship {
  return { id: readAttr(node, 'Id') ?? '', type: readAttr(node, 'Type') ?? '', target: readAttr(node, 'Target') ?? '' };
}
function relationshipToNode(rel: Relationship): XmlNode { return { '@_Id': rel.id, '@_Type': rel.type, '@_Target': rel.target }; }
