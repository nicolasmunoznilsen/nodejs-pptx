import type JSZip from 'jszip';
import { readAttr, walk, type XmlNode } from './xml.js';

export function nextShapeId(spTree: XmlNode): string {
  let max = 0;
  walk(spTree, node => {
    if (!Object.prototype.hasOwnProperty.call(node, 'p:cNvPr')) return;
    const id = parseInt(readAttr(node['p:cNvPr'], 'id') ?? '0', 10);
    if (Number.isFinite(id)) max = Math.max(max, id);
  });
  return String(max + 1);
}

export function allocateSlidePath(zip: JSZip): string {
  const used = new Set<number>();
  (zip as any).forEach((path: string, file: { dir?: boolean }) => {
    if (file.dir) return;
    const match = /^ppt\/slides\/slide(\d+)\.xml$/.exec(path);
    if (match) used.add(Number(match[1]));
  });
  let n = 1;
  while (used.has(n)) n += 1;
  return `ppt/slides/slide${n}.xml`;
}

export function maxNumericRelationshipId(rels: Array<{ id: string }>): number {
  return rels.reduce((max, rel) => Math.max(max, parseInt(rel.id.replace(/^rId/, ''), 10) || 0), 0);
}
