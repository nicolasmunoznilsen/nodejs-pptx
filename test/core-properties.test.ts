import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Presentation, resolveRootRelationshipTarget } from '../src/index.js';

const CORE_PROPS_TYPE = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const OFFICE_DOCUMENT_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const CORE_PROPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';

interface RelEntry { id: string; type: string; target: string }

function parseRels(xml: string): RelEntry[] {
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map(match => {
    const tag = match[0];
    const attr = (name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1] ?? '';
    return { id: attr('Id'), type: attr('Type'), target: attr('Target') };
  });
}

function countFiles(zip: JSZip, name: string): number {
  let count = 0;
  zip.forEach(relativePath => { if (relativePath === name) count++; });
  return count;
}

async function bufferOf(presentation: Presentation): Promise<Buffer> {
  return presentation.toBuffer();
}

describe('core-properties root relationship', () => {
  it('registers a unique core-properties relationship in _rels/.rels that resolves to docProps/core.xml', async () => {
    const buffer = await bufferOf(Presentation.create({ title: 'T', author: 'A' }));
    const zip = await JSZip.loadAsync(buffer);
    const relsFile = zip.file('_rels/.rels');
    expect(relsFile, '_rels/.rels must exist').not.toBeNull();
    const rels = parseRels(await relsFile!.async('string'));

    const coreRels = rels.filter(rel => rel.type === CORE_PROPS_TYPE);
    expect(coreRels).toHaveLength(1);
    expect(resolveRootRelationshipTarget(coreRels[0].target)).toBe('docProps/core.xml');
    expect(zip.file('docProps/core.xml'), 'docProps/core.xml part must exist').not.toBeNull();

    const officeRels = rels.filter(rel => rel.type === OFFICE_DOCUMENT_TYPE);
    expect(officeRels).toHaveLength(1);
    expect(resolveRootRelationshipTarget(officeRels[0].target)).toBe('ppt/presentation.xml');

    const ids = rels.map(rel => rel.id);
    expect(ids.every(id => id.length > 0)).toBe(true);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids).toContain(coreRels[0].id);
  });

  it('still declares the core-properties content type override', async () => {
    const buffer = await bufferOf(Presentation.create({ title: 'T', author: 'A' }));
    const zip = await JSZip.loadAsync(buffer);
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
    expect(contentTypes).toContain('PartName="/docProps/core.xml"');
    expect(contentTypes).toContain(`ContentType="${CORE_PROPS_CONTENT_TYPE}"`);
  });

  it('resolves title and author through the core-properties relationship on open', async () => {
    const buffer = await bufferOf(Presentation.create({ title: 'T', author: 'A' }));
    const opened = await Presentation.open(buffer);
    expect(opened.title).toBe('T');
    expect(opened.author).toBe('A');
  });

  it('round-trips title and author changes with exactly one docProps/core.xml part', async () => {
    const created = Presentation.create({ title: 'T', author: 'A' });
    created.title = 'Round Trip Title';
    created.author = 'Round Trip Author';
    const buffer = await bufferOf(created);

    const zip = await JSZip.loadAsync(buffer);
    expect(countFiles(zip, 'docProps/core.xml')).toBe(1);

    const reopened = await Presentation.open(buffer);
    expect(reopened.title).toBe('Round Trip Title');
    expect(reopened.author).toBe('Round Trip Author');

    const again = await bufferOf(reopened);
    const finalZip = await JSZip.loadAsync(again);
    expect(countFiles(finalZip, 'docProps/core.xml')).toBe(1);
    const finalOpened = await Presentation.open(again);
    expect(finalOpened.title).toBe('Round Trip Title');
    expect(finalOpened.author).toBe('Round Trip Author');
  });

  it('opens without throwing when the core-properties relationship is removed', async () => {
    const buffer = await bufferOf(Presentation.create({ title: 'T', author: 'A' }));
    const zip = await JSZip.loadAsync(buffer);
    const relsXml = await zip.file('_rels/.rels')!.async('string');
    const stripped = parseRels(relsXml).filter(rel => rel.type !== CORE_PROPS_TYPE);
    expect(stripped.some(rel => rel.type === OFFICE_DOCUMENT_TYPE)).toBe(true);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${stripped.map(rel => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"/>`).join('')}</Relationships>`);
    const brokenBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    const opened = await Presentation.open(brokenBuffer);
    expect(typeof opened.title).toBe('string');
    expect(typeof opened.author).toBe('string');
    expect(() => opened.slides.add()).not.toThrow();
    await expect(bufferOf(opened)).resolves.toBeInstanceOf(Buffer);
  });

  it('returns empty core properties when neither the relationship nor the part exists', async () => {
    const buffer = await bufferOf(Presentation.create({ title: 'T', author: 'A' }));
    const zip = await JSZip.loadAsync(buffer);
    const relsXml = await zip.file('_rels/.rels')!.async('string');
    const stripped = parseRels(relsXml).filter(rel => rel.type !== CORE_PROPS_TYPE);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${stripped.map(rel => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"/>`).join('')}</Relationships>`);
    zip.remove('docProps/core.xml');

    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    expect(opened.title).toBe('');
    expect(opened.author).toBe('');
    expect(() => { opened.title = 'New'; opened.author = 'New Author'; }).not.toThrow();
    await expect(bufferOf(opened)).resolves.toBeInstanceOf(Buffer);
  });

  it('resolves absolute and parent-relative core-properties targets to docProps/core.xml', async () => {
    expect(resolveRootRelationshipTarget('/docProps/core.xml')).toBe('docProps/core.xml');
    expect(resolveRootRelationshipTarget('../docProps/core.xml')).toBe('docProps/core.xml');

    for (const target of ['/docProps/core.xml', '../docProps/core.xml']) {
      const zip = await JSZip.loadAsync(await bufferOf(Presentation.create({ title: 'T', author: 'A' })));
      const relsXml = await zip.file('_rels/.rels')!.async('string');
      zip.file('_rels/.rels', relsXml.replace(
        new RegExp(`(Type="${CORE_PROPS_TYPE}" Target=")[^"]*(")`),
        `$1${target}$2`
      ));
      const rewritten = await zip.generateAsync({ type: 'nodebuffer' });
      const opened = await Presentation.open(rewritten);
      expect(opened.title, `target ${target}`).toBe('T');
      expect(opened.author, `target ${target}`).toBe('A');
      opened.title = 'Moved Title';
      const out = await JSZip.loadAsync(await bufferOf(opened));
      expect(countFiles(out, 'docProps/core.xml')).toBe(1);
      expect(await out.file('docProps/core.xml')!.async('string')).toContain('Moved Title');
    }
  });
});
