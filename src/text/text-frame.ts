import { arr, type XmlNode } from '../ooxml/xml.js';
import { assertAlignment } from './alignment.js';
import { Paragraph } from './paragraph.js';
import type { ParagraphAlignment } from './alignment.js';

export class TextFrame {
  constructor(private readonly txBody: XmlNode, private readonly markDirty: () => void) {}
  get paragraphs(): Paragraph[] { return arr(this.txBody['a:p'] as XmlNode | XmlNode[] | undefined).map(p => new Paragraph(p, this.markDirty)); }
  get text(): string { return this.paragraphs.map(p => p.text).join('\n'); }
  set text(value: string) {
    const lines = value.split('\n');
    this.txBody['a:p'] = lines.map(line => ({ 'a:r': { 'a:t': line } }));
    this.markDirty();
  }

  /**
   * Convenience accessor over the paragraphs: reading returns the alignment only when every
   * paragraph agrees (otherwise `undefined`), writing applies it to all of them.
   */
  get alignment(): ParagraphAlignment | undefined {
    const paragraphs = this.paragraphs;
    if (!paragraphs.length) return undefined;
    const first = paragraphs[0].alignment;
    return paragraphs.every(paragraph => paragraph.alignment === first) ? first : undefined;
  }
  set alignment(value: ParagraphAlignment | undefined) {
    const next = assertAlignment(value);
    const paragraphs = this.paragraphs;
    if (!paragraphs.length || paragraphs.every(paragraph => paragraph.alignment === next)) return;
    for (const paragraph of paragraphs) paragraph.alignment = next;
    this.markDirty();
  }
}
