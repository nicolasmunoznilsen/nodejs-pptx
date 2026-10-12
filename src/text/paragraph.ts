import { arr, type XmlNode } from '../ooxml/xml.js';
import { assertAlignment, readAlignment, writeAlignment, type ParagraphAlignment } from './alignment.js';
import {
  assertBullet,
  assertEmu,
  assertParagraphLevel,
  assertPoints,
  readBullet,
  readEmuAttribute,
  readParagraphLevel,
  readParagraphSpacing,
  writeBullet,
  writeEmuAttribute,
  writeParagraphLevel,
  writeParagraphSpacing,
  type ParagraphBullet,
  type ParagraphBulletInput,
} from './paragraph-format.js';
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

  /**
   * Outline level (`a:pPr/@lvl`), an integer from 0 to 8. `undefined` means "inherited": the
   * attribute is absent and the level comes from the placeholder or list style.
   */
  get level(): number | undefined { return readParagraphLevel(this.pPrNode); }
  set level(value: number | undefined) {
    const next = assertParagraphLevel(value);
    if (this.level === next) return;
    writeParagraphLevel(this.pPr(true), next);
    this.markDirty();
  }

  /**
   * Hanging or first-line indent (`a:pPr/@indent`) in integer EMU. Negative values pull the first
   * line to the left of `leftMargin`, which is how bullet paragraphs are usually built.
   * `undefined` means "inherited" and removes the attribute.
   */
  get indentation(): number | undefined { return readEmuAttribute(this.pPrNode, 'indent'); }
  set indentation(value: number | undefined) {
    const next = assertEmu(value, 'indentation');
    if (this.indentation === next) return;
    writeEmuAttribute(this.pPr(true), 'indent', next);
    this.markDirty();
  }

  /**
   * Left margin of the paragraph block (`a:pPr/@marL`) in integer EMU. `undefined` means
   * "inherited" and removes the attribute.
   */
  get leftMargin(): number | undefined { return readEmuAttribute(this.pPrNode, 'marL'); }
  set leftMargin(value: number | undefined) {
    const next = assertEmu(value, 'leftMargin');
    if (this.leftMargin === next) return;
    writeEmuAttribute(this.pPr(true), 'marL', next);
    this.markDirty();
  }

  /**
   * Space before the paragraph (`a:pPr/a:spcBef/a:spcPts`), expressed in typographic points
   * (`1 pt = 12700 EMU = 100 centipoints`, the unit stored in XML). Spacing that the document
   * expresses as a percentage (`a:spcPct`) is not modeled and reads back as `undefined` without
   * being modified. Assigning `undefined` removes `a:spcBef`, so the value is inherited again.
   */
  get spaceBefore(): number | undefined { return readParagraphSpacing(this.pPrNode, 'spaceBefore'); }
  set spaceBefore(value: number | undefined) {
    const next = assertPoints(value, 'spaceBefore');
    if (this.spaceBefore === next) return;
    writeParagraphSpacing(this.pPr(true), 'spaceBefore', next);
    this.markDirty();
  }

  /**
   * Space after the paragraph (`a:pPr/a:spcAft/a:spcPts`), in typographic points. Same semantics
   * as {@link Paragraph.spaceBefore}.
   */
  get spaceAfter(): number | undefined { return readParagraphSpacing(this.pPrNode, 'spaceAfter'); }
  set spaceAfter(value: number | undefined) {
    const next = assertPoints(value, 'spaceAfter');
    if (this.spaceAfter === next) return;
    writeParagraphSpacing(this.pPr(true), 'spaceAfter', next);
    this.markDirty();
  }

  /**
   * Bullet marker of the paragraph: `{ type: 'none' }` (`a:buNone`), `{ type: 'char', char }`
   * (`a:buChar`) or `{ type: 'autoNumber', autoNumber, number }` (`a:buAutoNum`). The setter also
   * accepts the bare kind strings `'none'`, `'char'`, and `'autoNumber'`. `undefined` means
   * "inherited" and removes every marker element, leaving sibling formatting such as `a:buFont`
   * untouched.
   */
  get bullet(): ParagraphBullet | undefined { return readBullet(this.pPrNode); }
  set bullet(value: ParagraphBulletInput | undefined) {
    const next = assertBullet(value);
    if (sameBullet(this.bullet, next)) return;
    writeBullet(this.pPr(true), next);
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

/** Compares two normalized bullet descriptors, treating missing optional fields as equal. */
function sameBullet(current: ParagraphBullet | undefined, next: ParagraphBullet | undefined): boolean {
  if (current === undefined || next === undefined) return current === next;
  return (
    current.type === next.type &&
    (current.type === 'char' ? current.char === next.char : true) &&
    (current.type === 'autoNumber' ? current.autoNumber === next.autoNumber && current.number === next.number : true)
  );
}
