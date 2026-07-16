import { XMLBuilder, XMLParser } from 'fast-xml-parser';

export type XmlNode = Record<string, unknown>;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', preserveOrder: false, trimValues: false });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', suppressEmptyNode: true, format: false });

export function parseXml(name: string, xml: string): XmlNode {
  try {
    const parsed = parser.parse(xml) as XmlNode;
    if (parsed['?xml']) delete parsed['?xml'];
    return parsed;
  } catch (e) {
    throw new Error(`Invalid XML in ${name}: ${(e as Error).message}`);
  }
}

export function xmlToString(xml: XmlNode): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + builder.build(xml);
}

export function readAttr(node: unknown, attr: string): string | undefined {
  return typeof node === 'object' && node !== null ? (node as XmlNode)[`@_${attr}`] as string | undefined : undefined;
}

export function writeAttr(node: XmlNode, attr: string, value: string | undefined): void {
  if (value === undefined) delete node[`@_${attr}`]; else node[`@_${attr}`] = value;
}

export function arr<T>(value: T | T[] | undefined): T[] { return value === undefined ? [] : Array.isArray(value) ? value : [value]; }

export function walk(value: unknown, visit: (node: XmlNode) => void): void {
  if (!value || typeof value !== 'object') return;
  if (Array.isArray(value)) { value.forEach(v => walk(v, visit)); return; }
  const node = value as XmlNode;
  visit(node);
  Object.values(node).forEach(v => walk(v, visit));
}
