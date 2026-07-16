import { arr, type XmlNode } from '../ooxml/xml.js';
import { Run } from './run.js';

export class Paragraph {
  constructor(private readonly node: XmlNode, private readonly markDirty: () => void) {}
  get runs(): Run[] { return arr(this.node['a:r'] as XmlNode | XmlNode[] | undefined).map(run => new Run(run, this.markDirty)); }
  get text(): string { return this.runs.map(run => run.text).join(''); }
  set text(value: string) {
    const runs = this.runs;
    if (!runs.length) {
      this.node['a:r'] = { 'a:rPr': {}, 'a:t': value };
    } else {
      runs[0].text = value;
      for (const run of runs.slice(1)) run.text = '';
    }
    this.markDirty();
  }
}
