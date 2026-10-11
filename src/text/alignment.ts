import { readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';

/**
 * Paragraph horizontal text alignment, mirroring the `ST_TextAlignType` values of the
 * `algn` attribute on `a:pPr` and the `PP_ALIGN` members used by `python-pptx`.
 */
export type ParagraphAlignment = 'left' | 'center' | 'right' | 'justify';

const ALIGNMENT_ENTRIES: ReadonlyArray<readonly [ParagraphAlignment, string]> = [
  ['left', 'l'],
  ['center', 'ctr'],
  ['right', 'r'],
  ['justify', 'just'],
];

/** Public names, in declaration order. */
export const ALIGNMENT_NAMES: readonly ParagraphAlignment[] = Object.freeze(
  ALIGNMENT_ENTRIES.map(([name]) => name),
);

/**
 * Builds a lookup map that carries no prototype. Inherited members such as `toString`,
 * `constructor`, `hasOwnProperty` or `__proto__` must never resolve to a value here,
 * otherwise untrusted input (or a user typo) would pass validation and leak a function
 * into the generated XML.
 */
function lookupTable<T>(entries: ReadonlyArray<readonly [string, T]>): Readonly<Record<string, T>> {
  const table: Record<string, T> = Object.create(null) as Record<string, T>;
  for (const [key, value] of entries) table[key] = value;
  return Object.freeze(table);
}

/** Public name -> OOXML `algn` value. Own properties only: inherited names are not alignments. */
export const ALIGNMENT_TO_OOXML = lookupTable<string>(ALIGNMENT_ENTRIES) as Readonly<Record<ParagraphAlignment, string>>;

/** OOXML `algn` value -> public name. Own properties only, for the same reason. */
export const OOXML_TO_ALIGNMENT = lookupTable<ParagraphAlignment>(
  ALIGNMENT_ENTRIES.map(([name, value]) => [value, name]),
) as Readonly<Record<string, ParagraphAlignment>>;

/** True only for the supported public alignment names; inherited property names are rejected. */
export function isAlignment(value: unknown): value is ParagraphAlignment {
  return typeof value === 'string' && Object.prototype.hasOwnProperty.call(ALIGNMENT_TO_OOXML, value);
}

/**
 * Reads the `algn` attribute of an `a:pPr` node. Returns `undefined` when the alignment is
 * not explicit, which means it is inherited from the paragraph/list style or the placeholder,
 * and also when the value is not a known `ST_TextAlignType` member.
 */
export function readAlignment(pPr: XmlNode | undefined): ParagraphAlignment | undefined {
  const raw = readAttr(pPr, 'algn');
  if (raw === undefined) return undefined;
  return Object.prototype.hasOwnProperty.call(OOXML_TO_ALIGNMENT, raw) ? OOXML_TO_ALIGNMENT[raw] : undefined;
}

/** Writes or clears (`undefined`) the `algn` attribute of an `a:pPr` node. */
export function writeAlignment(pPr: XmlNode, value: ParagraphAlignment | undefined): void {
  if (value === undefined) {
    writeAttr(pPr, 'algn', undefined);
    return;
  }
  if (!isAlignment(value)) throw invalidAlignment(value);
  writeAttr(pPr, 'algn', ALIGNMENT_TO_OOXML[value]);
}

/** Throws a descriptive error when `value` is not a supported alignment. */
export function assertAlignment(value: unknown): ParagraphAlignment | undefined {
  if (value === undefined || value === null) return undefined;
  if (!isAlignment(value)) throw invalidAlignment(value);
  return value;
}

function invalidAlignment(value: unknown): TypeError {
  return new TypeError(`Invalid paragraph alignment: ${JSON.stringify(value)}. Expected one of: ${ALIGNMENT_NAMES.join(', ')}, or undefined.`);
}
