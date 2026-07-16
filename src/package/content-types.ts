import JSZip from 'jszip';

export class ContentTypes {
  dirty = false;
  constructor(private xml: string) {}

  static create(): ContentTypes {
    return new ContentTypes('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
  }

  static async load(zip: JSZip): Promise<ContentTypes> {
    const file = zip.file('[Content_Types].xml');
    if (!file) throw new Error('Invalid PPTX: missing [Content_Types].xml.');
    return new ContentTypes(await file.async('string'));
  }

  hasOverride(partName: string): boolean { return this.xml.includes(`PartName="/${trim(partName)}"`); }

  addOverride(partName: string, contentType: string): void {
    const part = trim(partName);
    if (this.hasOverride(part)) return;
    this.xml = this.xml.replace('</Types>', `<Override PartName="/${part}" ContentType="${contentType}"/></Types>`);
    this.dirty = true;
  }

  removeOverride(partName: string): void {
    const part = trim(partName).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const before = this.xml;
    this.xml = this.xml.replace(new RegExp(`<Override\\s+PartName="/${part}"\\s+ContentType="[^"]+"\\s*/>`), '');
    if (this.xml !== before) this.dirty = true;
  }

  save(zip: JSZip, force = false): void { if (force || this.dirty) zip.file('[Content_Types].xml', this.xml); }
  toString(): string { return this.xml; }
}

function trim(partName: string): string { return partName.startsWith('/') ? partName.slice(1) : partName; }
