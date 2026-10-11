import { readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';

/**
 * Paragraph horizontal text alignment, mirroring the `ST_TextAlignType` values of the
 * `algn` attribute on `a:pPr` and the `PP_ALIGN` members used by `python-pptx`.
 */
export type ParagraphAlignment = 'left' | 'center' | 'right' | 'justify';

/** Public name -> OOXML `algn` value. Frozen so the lookup tables cannot be mutated at runtime. */
export const ALIGNMENT_TO_OOXML: Readonly<Record<ParagraphAlignment, string>> = Object.freeze({
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
});

/** OOXML `algn` value -> public name. */
export const OOXML_TO_ALIGNMENT: Readonly<Record<string, ParagraphAlignment>> = Object.freeze(
  Object.fromEntries(
    Object.entries(ALIGNMENT_TO_OOXML).map(([name, value]) => [value, name as ParagraphAlignment]),
  ) as Record<string, ParagraphAlignment>,
);

/**
 * Public alignment name -> OOXML `algn` value, resolved through an own-property lookup.
 *
 * The guard is `Object.hasOwn`, not `in` and not a truthiness check, so names inherited from
 * `Object.prototype` (`constructor`, `toString`, `valueOf`, `__proto__`, ...) never resolve
 * through the prototype chain and never leak an inherited value into generated XML.
 */
export function alignmentToOoxml(value: unknown): string | undefined {
  if (typeof value !== 'string' || !Object.hasOwn(ALIGNMENT_TO_OOXML, value)) return undefined;
  return (ALIGNMENT_TO_OOXML as Record<string, string>)[value];
}

/**
 * OOXML `algn` value -> public alignment name, resolved through an own-property lookup.
 *
 * Unknown, malformed, and prototype-inherited values (`constructor`, `__proto__`, ...) resolve to
 * `undefined`, which the public API reports as "inherited alignment" instead of handing a
 * non-alignment value (for example the `Object` constructor) back to callers.
 */
export function alignmentFromOoxml(raw: unknown): ParagraphAlignment | undefined {
  if (typeof raw !== 'string' || !Object.hasOwn(OOXML_TO_ALIGNMENT, raw)) return undefined;
  return OOXML_TO_ALIGNMENT[raw];
}

/**
 * Reads the `algn` attribute of an `a:pPr` node. Returns `undefined` when the alignment is
 * not explicit or carries a value this library does not understand, which means it is inherited
 * from the paragraph/list style or the placeholder.
 */
export function readAlignment(pPr: XmlNode | undefined): ParagraphAlignment | undefined {
  return alignmentFromOoxml(readAttr(pPr, 'algn'));
}

/** Writes or clears (`undefined`) the `algn` attribute of an `a:pPr` node. */
export function writeAlignment(pPr: XmlNode, value: ParagraphAlignment | undefined): void {
  if (value === undefined) {
    writeAttr(pPr, 'algn', undefined);
    return;
  }
  const raw = alignmentToOoxml(value);
  if (raw === undefined) throw new TypeError(describeInvalidAlignment(value));
  writeAttr(pPr, 'algn', raw);
}

/** @internal Shared message for rejected alignment values. */
function describeInvalidAlignment(value: unknown): string {
  return `Invalid paragraph alignment: ${JSON.stringify(value)}. Expected one of: ${Object.keys(ALIGNMENT_TO_OOXML).join(', ')}, or undefined.`;
}

/** Throws a descriptive error when `value` is not a supported alignment. */
export function assertAlignment(value: unknown): ParagraphAlignment | undefined {
  if (value === undefined || value === null) return undefined;
  if (alignmentToOoxml(value) === undefined) {
    throw new TypeError(describeInvalidAlignment(value));
  }
  return value as ParagraphAlignment;
}
