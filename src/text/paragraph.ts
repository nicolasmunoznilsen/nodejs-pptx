import { arr, type XmlNode } from '../ooxml/xml.js';
import { assertAlignment, readAlignment, writeAlignment, type ParagraphAlignment } from './alignment.js';
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

  /**
   * Horizontal alignment of this paragraph. `undefined` means "inherited": the `a:pPr` node
   * carries no explicit `algn` attribute, so PowerPoint resolves it through the paragraph
   * list style, the placeholder, or the theme.
   */
  get alignment(): ParagraphAlignment | undefined { return readAlignment(this.pPrNode); }
  set alignment(value: ParagraphAlignment | undefined) {
    const next = assertAlignment(value);
    if (this.alignment === next) return;
    writeAlignment(this.pPr(true), next);
    this.markDirty();
  }

  /** @internal The existing `a:pPr` node, without creating one. */
  private get pPrNode(): XmlNode | undefined {
    const pPr = this.node['a:pPr'];
    if (Array.isArray(pPr)) return pPr[0] as XmlNode | undefined;
    return pPr && typeof pPr === 'object' ? (pPr as XmlNode) : undefined;
  }

  /**
   * Returns the `a:pPr` node, creating it as the first child of `a:p` when missing. The schema
   * requires `a:pPr` to precede every `a:r` and `a:endParaRPr`, so a freshly created node is
   * re-inserted at the front instead of being appended to the end of the object.
   */
  private pPr(_create: true): XmlNode {
    const existing = this.pPrNode;
    if (existing) return existing;
    const pPr: XmlNode = {};
    const rest = { ...this.node };
    for (const key of Object.keys(this.node)) delete this.node[key];
    this.node['a:pPr'] = pPr;
    for (const [key, value] of Object.entries(rest)) {
      if (key !== 'a:pPr') this.node[key] = value;
    }
    return pPr;
  }
}
