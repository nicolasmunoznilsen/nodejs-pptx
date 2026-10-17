import { arr, type XmlNode } from '../ooxml/xml.js';
import { assertAlignment } from './alignment.js';
import { assertMarginInches, explicitMarginEmu, MARGIN_DEFAULT_EMU, marginEmu, readMarginInches, writeMarginInches, type TextFrameMarginSide } from './margins.js';
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

  /**
   * Left text-frame inset in inches, backed by `a:bodyPr/@lIns`. `undefined` means the attribute
   * is absent, so PowerPoint applies its 0.1" default; the getter reports that default instead of
   * `undefined`, because it is what is rendered.
   */
  get marginLeft(): number { return readMarginInches(this.bodyPrNode, 'left'); }
  set marginLeft(value: number | undefined) { this.writeMargin('left', value); }
  get marginTop(): number { return readMarginInches(this.bodyPrNode, 'top'); }
  set marginTop(value: number | undefined) { this.writeMargin('top', value); }
  get marginRight(): number { return readMarginInches(this.bodyPrNode, 'right'); }
  set marginRight(value: number | undefined) { this.writeMargin('right', value); }
  get marginBottom(): number { return readMarginInches(this.bodyPrNode, 'bottom'); }
  set marginBottom(value: number | undefined) { this.writeMargin('bottom', value); }

  /**
   * Convenience accessor over the four insets: reading returns the shared value only when all four
   * agree (otherwise `undefined`), writing applies it to all of them.
   */
  get margin(): number | undefined {
    const left = this.marginLeft;
    return [this.marginTop, this.marginRight, this.marginBottom].every(value => value === left)
      ? left
      : undefined;
  }
  set margin(value: number | undefined) {
    const next = assertMarginInches(value);
    const changed = (['left', 'top', 'right', 'bottom'] as const).filter(side => !this.isSameMargin(side, next));
    if (!changed.length) return;
    const bodyPr = this.bodyPr(true);
    for (const side of changed) writeMarginInches(bodyPr, side, next);
    this.markDirty();
  }

  /** @internal Single-side write: validates, skips no-op writes, and only then marks dirty. */
  private writeMargin(side: TextFrameMarginSide, value: number | undefined): void {
    const next = assertMarginInches(value);
    if (this.isSameMargin(side, next)) return;
    writeMarginInches(this.bodyPr(true), side, next);
    this.markDirty();
  }

  /**
   * @internal Whether `next` is already the effective EMU value of `side`.
   *
   * The comparison uses the effective inset, i.e. the explicit `@_*Ins` attribute when present and
   * otherwise the OOXML default for that side (`explicitMarginEmu(...) ?? MARGIN_DEFAULT_EMU[side]`),
   * so re-assigning an inherited default (`frame.marginLeft = 0.1` on a document without `lIns`) is a
   * no-op instead of an edit that marks the slide dirty and inserts `lIns="91440"`.
   *
   * `undefined` is deliberately compared against the explicit attribute only: it means "inherit", so
   * an explicit attribute that happens to equal the default is still removed.
   */
  private isSameMargin(side: TextFrameMarginSide, next: number | undefined): boolean {
    const current = explicitMarginEmu(this.bodyPrNode, side);
    if (next === undefined) return current === undefined;
    const effective = current ?? MARGIN_DEFAULT_EMU[side];
    return effective === marginEmu(next);
  }

  /** @internal The existing `a:bodyPr` node, without creating one. */
  private get bodyPrNode(): XmlNode | undefined {
    const bodyPr = this.txBody['a:bodyPr'];
    if (Array.isArray(bodyPr)) return bodyPr[0] as XmlNode | undefined;
    return bodyPr && typeof bodyPr === 'object' ? (bodyPr as XmlNode) : undefined;
  }

  /**
   * Returns the `a:bodyPr` node, creating it as the first child of `p:txBody` when missing. The
   * schema requires `a:bodyPr` to precede `a:lstStyle` and every `a:p`, so a freshly created node
   * is re-inserted at the front instead of being appended to the end of the object.
   */
  private bodyPr(_create: true): XmlNode {
    const existing = this.bodyPrNode;
    if (existing) return existing;
    const bodyPr: XmlNode = {};
    const rest = { ...this.txBody };
    for (const key of Object.keys(this.txBody)) delete this.txBody[key];
    this.txBody['a:bodyPr'] = bodyPr;
    for (const [key, value] of Object.entries(rest)) {
      if (key !== 'a:bodyPr') this.txBody[key] = value;
    }
    return bodyPr;
  }
}
