import { readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';

/**
 * Paragraph horizontal text alignment, mirroring the `ST_TextAlignType` values of the
 * `algn` attribute on `a:pPr` and the `PP_ALIGN` members used by `python-pptx`.
 */
export type ParagraphAlignment = 'left' | 'center' | 'right' | 'justify';

/** Public name -> OOXML `algn` value. */
export const ALIGNMENT_TO_OOXML: Readonly<Record<ParagraphAlignment, string>> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
};

/** OOXML `algn` value -> public name. */
export const OOXML_TO_ALIGNMENT: Readonly<Record<string, ParagraphAlignment>> = Object.fromEntries(
  Object.entries(ALIGNMENT_TO_OOXML).map(([name, value]) => [value, name as ParagraphAlignment]),
) as Record<string, ParagraphAlignment>;

/**
 * Reads the `algn` attribute of an `a:pPr` node. Returns `undefined` when the alignment is
 * not explicit, which means it is inherited from the paragraph/list style or the placeholder.
 */
export function readAlignment(pPr: XmlNode | undefined): ParagraphAlignment | undefined {
  const raw = readAttr(pPr, 'algn');
  if (raw === undefined) return undefined;
  return OOXML_TO_ALIGNMENT[raw];
}

/** Writes or clears (`undefined`) the `algn` attribute of an `a:pPr` node. */
export function writeAlignment(pPr: XmlNode, value: ParagraphAlignment | undefined): void {
  if (value === undefined) {
    writeAttr(pPr, 'algn', undefined);
    return;
  }
  writeAttr(pPr, 'algn', ALIGNMENT_TO_OOXML[value]);
}

/** Throws a descriptive error when `value` is not a supported alignment. */
export function assertAlignment(value: unknown): ParagraphAlignment | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string' || !(value in ALIGNMENT_TO_OOXML)) {
    throw new TypeError(`Invalid paragraph alignment: ${JSON.stringify(value)}. Expected one of: ${Object.keys(ALIGNMENT_TO_OOXML).join(', ')}, or undefined.`);
  }
  return value as ParagraphAlignment;
}
