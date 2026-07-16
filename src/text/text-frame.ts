import { arr, type XmlNode } from '../ooxml/xml.js';
import { Paragraph } from './paragraph.js';

export class TextFrame {
  constructor(private readonly txBody: XmlNode, private readonly markDirty: () => void) {}
  get paragraphs(): Paragraph[] { return arr(this.txBody['a:p'] as XmlNode | XmlNode[] | undefined).map(p => new Paragraph(p, this.markDirty)); }
  get text(): string { return this.paragraphs.map(p => p.text).join('\n'); }
}
