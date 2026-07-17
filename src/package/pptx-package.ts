import JSZip from 'jszip';

export class PptxPackage {
  constructor(readonly zip: JSZip) {}
  static create(): PptxPackage { return new PptxPackage(new JSZip()); }
  static async open(data: Buffer): Promise<PptxPackage> { return new PptxPackage(await JSZip.loadAsync(data)); }
  file(path: string): unknown { return this.zip.file(path); }
  write(path: string, data: string | Buffer): void { this.zip.file(path, data); }
}
