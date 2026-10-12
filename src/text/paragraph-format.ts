import { arr, readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';

/**
 * Paragraph formatting helpers: outline level, indentation/margins, spacing before/after, and
 * bullets. Like `alignment.ts`, every public mapping lives in a frozen lookup table and every
 * setter path goes through an `assert*` function that throws a descriptive `TypeError`.
 *
 * All accessors treat `undefined` as "inherited": the corresponding attribute or child of
 * `a:pPr` is removed, so PowerPoint resolves the value through the list style, the placeholder,
 * or the theme.
 */

/** `a:pPr/@lvl` is an unsigned byte restricted to the nine outline levels of DrawingML. */
export const MIN_PARAGRAPH_LEVEL = 0;
export const MAX_PARAGRAPH_LEVEL = 8;

/** One typographic point in English Metric Units (`1 pt = 12700 EMU`). */
export const EMU_PER_POINT = 12700;

/** Largest value the OOXML `int` attributes used here can carry. */
const MAX_INT32 = 2147483647;

/**
 * Child order required by `CT_TextParagraphProperties` (`a:pPr`). Children not listed here keep
 * their relative order and are appended after the known elements, so unknown or future elements
 * survive a round trip.
 */
export const PARAGRAPH_PROPERTY_ORDER: readonly string[] = Object.freeze([
  'a:lnSpc',
  'a:spcBef',
  'a:spcAft',
  'a:buClrTx',
  'a:buClr',
  'a:buSzTx',
  'a:buSzPct',
  'a:buSzPts',
  'a:buFontTx',
  'a:buFont',
  'a:buNone',
  'a:buAutoNum',
  'a:buChar',
  'a:tabLst',
  'a:defRPr',
  'a:extLst',
]);

/** The three mutually exclusive bullet markers of `CT_TextParagraphProperties`. */
const BULLET_ELEMENTS: readonly string[] = Object.freeze(['a:buNone', 'a:buAutoNum', 'a:buChar']);

/** Bullet kind, mirroring the `a:buNone` / `a:buChar` / `a:buAutoNum` children of `a:pPr`. */
export type ParagraphBulletType = 'none' | 'char' | 'autoNumber';

/**
 * Supported `ST_TextAutoNumber` values for `a:buAutoNum/@type`. Values outside this subset are
 * preserved in the XML as long as they are not rewritten, and are reported without an
 * `autoNumber` field.
 */
export type ParagraphAutoNumberType =
  | 'arabicPeriod'
  | 'arabicParen'
  | 'arabicParenR'
  | 'arabicComma'
  | 'arabicDash'
  | 'alphaLcPeriod'
  | 'alphaLcParen'
  | 'alphaLcLParen'
  | 'alphaUcPeriod'
  | 'alphaUcParen'
  | 'alphaUcLParen'
  | 'nonNumLcAlphPeriod'
  | 'nonNumLcAlphParen'
  | 'nonNumUcAlphPeriod'
  | 'nonNumUcAlphParen'
  | 'doubleDash'
  | 'hyphen'
  | 'ellipseChsDbPeriod'
  | 'circleNumTxtChsSimple'
  | 'circleNumDbChsSimple'
  | 'circleNumTxtChsHanging'
  | 'circleNumDbChsHanging'
  | 'circleNumDbChsFullWidth';

/** Public name -> OOXML `a:buAutoNum/@type` value. Frozen so it cannot be mutated at runtime. */
export const AUTO_NUMBER_TO_OOXML: Readonly<Record<ParagraphAutoNumberType, string>> = Object.freeze({
  arabicPeriod: 'arabicPeriod',
  arabicParen: 'arabicParen',
  arabicParenR: 'arabicParenR',
  arabicComma: 'arabicComma',
  arabicDash: 'arabicDash',
  alphaLcPeriod: 'alphaLcPeriod',
  alphaLcParen: 'alphaLcParen',
  alphaLcLParen: 'alphaLcLParen',
  alphaUcPeriod: 'alphaUcPeriod',
  alphaUcParen: 'alphaUcParen',
  alphaUcLParen: 'alphaUcLParen',
  nonNumLcAlphPeriod: 'nonNumLcAlphPeriod',
  nonNumLcAlphParen: 'nonNumLcAlphParen',
  nonNumUcAlphPeriod: 'nonNumUcAlphPeriod',
  nonNumUcAlphParen: 'nonNumUcAlphParen',
  doubleDash: 'doubleDash',
  hyphen: 'hyphen',
  ellipseChsDbPeriod: 'ellipseChsDbPeriod',
  circleNumTxtChsSimple: 'circleNumTxtChsSimple',
  circleNumDbChsSimple: 'circleNumDbChsSimple',
  circleNumTxtChsHanging: 'circleNumTxtChsHanging',
  circleNumDbChsHanging: 'circleNumDbChsHanging',
  circleNumDbChsFullWidth: 'circleNumDbChsFullWidth',
});

/** OOXML `a:buAutoNum/@type` value -> public name. */
export const OOXML_TO_AUTO_NUMBER: Readonly<Record<string, ParagraphAutoNumberType>> = Object.freeze(
  Object.fromEntries(
    Object.entries(AUTO_NUMBER_TO_OOXML).map(([name, value]) => [value, name as ParagraphAutoNumberType]),
  ) as Record<string, ParagraphAutoNumberType>,
);

/** Auto-number name -> OOXML `type` value, resolved through an own-property lookup. */
export function autoNumberToOoxml(value: unknown): string | undefined {
  if (typeof value !== 'string' || !Object.hasOwn(AUTO_NUMBER_TO_OOXML, value)) return undefined;
  return (AUTO_NUMBER_TO_OOXML as Record<string, string>)[value];
}

/** OOXML `type` value -> auto-number name. Prototype-inherited names resolve to `undefined`. */
export function autoNumberFromOoxml(raw: unknown): ParagraphAutoNumberType | undefined {
  if (typeof raw !== 'string' || !Object.hasOwn(OOXML_TO_AUTO_NUMBER, raw)) return undefined;
  return OOXML_TO_AUTO_NUMBER[raw];
}

/** Default `a:buAutoNum/@type` used when a caller asks for numbering without naming a scheme. */
export const DEFAULT_AUTO_NUMBER_TYPE: ParagraphAutoNumberType = 'arabicPeriod';

/** A paragraph bullet: the marker kind plus its character or numbering details. */
export interface ParagraphBullet {
  /** Marker kind. `'none'` writes `a:buNone`, so the paragraph is explicitly not bulleted. */
  type: ParagraphBulletType;
  /** Bullet glyph for `type: 'char'`, written to `a:buChar/@char`. */
  char?: string;
  /** Numbering scheme for `type: 'autoNumber'`, written to `a:buAutoNum/@type`. */
  autoNumber?: ParagraphAutoNumberType;
  /** Restart value for `type: 'autoNumber'`, written to `a:buAutoNum/@num`. */
  number?: number;
}

/** Accepted values of `paragraph.bullet`: a bare kind or a full descriptor. */
export type ParagraphBulletInput = ParagraphBulletType | ParagraphBullet;

/** Which spacing a helper refers to: `a:spcBef` or `a:spcAft`. */
export type ParagraphSpacingKind = 'spaceBefore' | 'spaceAfter';

const SPACING_ELEMENT: Readonly<Record<ParagraphSpacingKind, string>> = Object.freeze({
  spaceBefore: 'a:spcBef',
  spaceAfter: 'a:spcAft',
});

/** Points -> `a:spcPts/@val` centipoints. */
export function pointsToCentipoints(points: number): number {
  return Math.round(points * 100);
}

/** `a:spcPts/@val` centipoints -> points. */
export function centipointsToPoints(centipoints: number): number {
  return centipoints / 100;
}

/** Points -> EMU, for callers that prefer to think in English Metric Units. */
export function pointsToEmu(points: number): number {
  return Math.round(points * EMU_PER_POINT);
}

/** EMU -> points, the inverse of {@link pointsToEmu}. */
export function emuToPoints(emu: number): number {
  return emu / EMU_PER_POINT;
}

/** Throws a descriptive `TypeError` unless `value` is an integer outline level, or `undefined`. */
export function assertParagraphLevel(value: unknown): number | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'number' || !Number.isInteger(value) || value < MIN_PARAGRAPH_LEVEL || value > MAX_PARAGRAPH_LEVEL) {
    throw new TypeError(
      `Invalid paragraph level: ${describe(value)}. Expected an integer between ${MIN_PARAGRAPH_LEVEL} and ${MAX_PARAGRAPH_LEVEL}, or undefined.`,
    );
  }
  return value;
}

/** Throws a descriptive `TypeError` unless `value` is an integer EMU length, or `undefined`. */
export function assertEmu(value: unknown, name: string): number | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'number' || !Number.isInteger(value) || !Number.isSafeInteger(value) || Math.abs(value) > MAX_INT32) {
    throw new TypeError(`Invalid paragraph ${name}: ${describe(value)}. Expected an integer EMU value, or undefined.`);
  }
  return value;
}

/**
 * Throws a descriptive `TypeError` unless `value` is a finite point distance, or `undefined`.
 * The returned value is rounded to the centipoint resolution the XML can actually store.
 */
export function assertPoints(value: unknown, name: string): number | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Invalid paragraph ${name}: ${describe(value)}. Expected a finite number of points, or undefined.`);
  }
  const centipoints = pointsToCentipoints(value);
  if (Math.abs(centipoints) > MAX_INT32) {
    throw new TypeError(`Invalid paragraph ${name}: ${describe(value)}. The value exceeds the OOXML range.`);
  }
  return centipointsToPoints(centipoints);
}

/** Throws a descriptive `TypeError` unless `value` is a supported bullet descriptor. */
export function assertBullet(value: unknown): ParagraphBullet | undefined {
  if (value === undefined || value === null) return undefined;
  const kind = typeof value === 'string' ? { type: value } : value;
  if (typeof kind !== 'object' || kind === null) throw new TypeError(describeInvalidBullet(value));
  const candidate = kind as Partial<ParagraphBullet>;
  if (
    typeof candidate.type !== 'string' ||
    !Object.hasOwn(BULLET_ELEMENT_BY_TYPE as unknown as Record<string, string>, candidate.type)
  ) {
    throw new TypeError(describeInvalidBullet(value));
  }
  const bullet: ParagraphBullet = { type: candidate.type as ParagraphBulletType };
  if (candidate.char !== undefined && candidate.char !== null) {
    if (typeof candidate.char !== 'string' || candidate.char === '') {
      throw new TypeError(`Invalid bullet char: ${describe(candidate.char)}. Expected a non-empty string.`);
    }
    bullet.char = candidate.char;
  }
  if (candidate.autoNumber !== undefined && candidate.autoNumber !== null) {
    if (autoNumberToOoxml(candidate.autoNumber) === undefined) {
      throw new TypeError(
        `Invalid bullet autoNumber type: ${describe(candidate.autoNumber)}. Expected one of: ${Object.keys(AUTO_NUMBER_TO_OOXML).join(', ')}.`,
      );
    }
    bullet.autoNumber = candidate.autoNumber;
  }
  if (candidate.number !== undefined && candidate.number !== null) {
    if (typeof candidate.number !== 'number' || !Number.isInteger(candidate.number) || candidate.number < 1 || candidate.number > MAX_INT32) {
      throw new TypeError(`Invalid bullet number: ${describe(candidate.number)}. Expected a positive integer, or undefined.`);
    }
    bullet.number = candidate.number;
  }
  if (bullet.type === 'char' && bullet.char === undefined) bullet.char = '•';
  if (bullet.type === 'char' && bullet.autoNumber !== undefined) delete bullet.autoNumber;
  if (bullet.type === 'autoNumber' && bullet.autoNumber === undefined) bullet.autoNumber = DEFAULT_AUTO_NUMBER_TYPE;
  if (bullet.type === 'none') {
    delete bullet.char;
    delete bullet.autoNumber;
    delete bullet.number;
  }
  return bullet;
}

const BULLET_ELEMENT_BY_TYPE: Readonly<Record<ParagraphBulletType, string>> = Object.freeze({
  none: 'a:buNone',
  char: 'a:buChar',
  autoNumber: 'a:buAutoNum',
});

/**
 * Reorders the children of an `a:pPr` node to match {@link PARAGRAPH_PROPERTY_ORDER}. Attributes
 * and every child value are kept untouched; unknown children keep their relative order at the end.
 */
export function orderParagraphProperties(pPr: XmlNode): void {
  const entries = Object.entries(pPr);
  const attributes = entries.filter(([key]) => key.startsWith('@_'));
  const children = entries.filter(([key]) => !key.startsWith('@_'));
  if (!children.length) return;
  const rank = (key: string): number => {
    const index = PARAGRAPH_PROPERTY_ORDER.indexOf(key);
    return index === -1 ? PARAGRAPH_PROPERTY_ORDER.length : index;
  };
  const ordered = children
    .map((entry, index) => ({ entry, index, order: rank(entry[0]) }))
    .sort((a, b) => a.order - b.order || a.index - b.index)
    .map(item => item.entry);
  for (const key of Object.keys(pPr)) delete pPr[key];
  for (const [key, value] of attributes) pPr[key] = value;
  for (const [key, value] of ordered) pPr[key] = value;
}

/** Reads `a:pPr/@lvl`. Non-numeric or out-of-range values read back as inherited. */
export function readParagraphLevel(pPr: XmlNode | undefined): number | undefined {
  const raw = typeof readAttr(pPr, 'lvl') === 'string' ? (readAttr(pPr, 'lvl') as string).trim() : undefined;
  if (raw === undefined || !/^\d+$/.test(raw)) return undefined;
  const value = Number(raw);
  return value >= MIN_PARAGRAPH_LEVEL && value <= MAX_PARAGRAPH_LEVEL ? value : undefined;
}

/** Writes or clears (`undefined`) the `lvl` attribute of an `a:pPr` node. */
export function writeParagraphLevel(pPr: XmlNode, value: number | undefined): void {
  const next = assertParagraphLevel(value);
  writeAttr(pPr, 'lvl', next === undefined ? undefined : String(next));
}

/** Reads an integer EMU attribute (`marL`, `indent`). Malformed values read back as inherited. */
export function readEmuAttribute(pPr: XmlNode | undefined, attribute: string): number | undefined {
  const raw = readAttr(pPr, attribute);
  if (raw === undefined) return undefined;
  if (typeof raw !== 'string' || !/^-?\d+$/.test(raw.trim())) return undefined;
  return Number(raw.trim());
}

/** Writes or clears (`undefined`) an integer EMU attribute (`marL`, `indent`). */
export function writeEmuAttribute(pPr: XmlNode, attribute: string, value: number | undefined): void {
  const next = assertEmu(value, attribute);
  writeAttr(pPr, attribute, next === undefined ? undefined : String(next));
}

/**
 * Reads `a:spcBef`/`a:spcAft` as points. Spacing expressed as a percentage (`a:spcPct`) is not
 * modeled and reads back as `undefined` while remaining untouched in the XML.
 */
export function readParagraphSpacing(pPr: XmlNode | undefined, kind: ParagraphSpacingKind): number | undefined {
  const container = arr(pPr?.[SPACING_ELEMENT[kind]] as XmlNode | XmlNode[] | undefined)[0] as XmlNode | undefined;
  if (!container || typeof container !== 'object') return undefined;
  const points = arr(container['a:spcPts'] as XmlNode | XmlNode[] | undefined)[0] as XmlNode | undefined;
  const raw = readAttr(points, 'val');
  if (typeof raw !== 'string' || !/^-?\d+$/.test(raw.trim())) return undefined;
  return centipointsToPoints(Number(raw.trim()));
}

/** Writes or clears (`undefined`) `a:spcBef`/`a:spcAft` using `a:spcPts`, then reorders `a:pPr`. */
export function writeParagraphSpacing(pPr: XmlNode, kind: ParagraphSpacingKind, value: number | undefined): void {
  const element = SPACING_ELEMENT[kind];
  if (value === undefined) {
    delete pPr[element];
    return;
  }
  const points = assertPoints(value, kind);
  if (points === undefined) return;
  pPr[element] = { 'a:spcPts': { '@_val': String(pointsToCentipoints(points)) } };
  orderParagraphProperties(pPr);
}

/** Reads the bullet marker of an `a:pPr` node, or `undefined` when it is inherited. */
export function readBullet(pPr: XmlNode | undefined): ParagraphBullet | undefined {
  if (!pPr) return undefined;
  if (childExists(pPr, 'a:buNone')) return { type: 'none' };
  if (childExists(pPr, 'a:buChar')) {
    const char = readAttr(childOf(pPr, 'a:buChar'), 'char');
    return { type: 'char', char: char === undefined ? '' : char };
  }
  if (childExists(pPr, 'a:buAutoNum')) {
    const node = childOf(pPr, 'a:buAutoNum');
    const bullet: ParagraphBullet = { type: 'autoNumber', autoNumber: autoNumberFromOoxml(readAttr(node, 'type')) ?? DEFAULT_AUTO_NUMBER_TYPE };
    const raw = readAttr(node, 'num');
    if (typeof raw === 'string' && /^\d+$/.test(raw.trim())) bullet.number = Number(raw.trim());
    return bullet;
  }
  return undefined;
}

/**
 * Writes or clears (`undefined`) the bullet marker of an `a:pPr` node. Only the three marker
 * elements are replaced; sibling formatting such as `a:buClr`, `a:buSzPct`, and `a:buFont` is
 * preserved, and the children are reordered to satisfy the schema.
 */
export function writeBullet(pPr: XmlNode, value: ParagraphBulletInput | undefined): void {
  const bullet = assertBullet(value);
  for (const element of BULLET_ELEMENTS) delete pPr[element];
  if (bullet) {
    const element = BULLET_ELEMENT_BY_TYPE[bullet.type];
    if (bullet.type === 'char') pPr[element] = { '@_char': bullet.char ?? '•' };
    else if (bullet.type === 'autoNumber') {
      const node: XmlNode = { '@_type': autoNumberToOoxml(bullet.autoNumber ?? DEFAULT_AUTO_NUMBER_TYPE) };
      if (bullet.number !== undefined) node['@_num'] = String(bullet.number);
      pPr[element] = node;
    } else pPr[element] = {};
  }
  orderParagraphProperties(pPr);
}

function childOf(pPr: XmlNode, key: string): XmlNode | undefined {
  return arr(pPr[key] as XmlNode | XmlNode[] | undefined)[0] as XmlNode | undefined;
}

function childExists(pPr: XmlNode, key: string): boolean {
  const value = pPr[key];
  return value !== undefined && arr(value as XmlNode | XmlNode[]).length > 0;
}

function describe(value: unknown): string {
  return typeof value === 'bigint' ? `${value}n` : JSON.stringify(value) ?? String(value);
}

function describeInvalidBullet(value: unknown): string {
  return `Invalid paragraph bullet: ${describe(value)}. Expected 'none', 'char', 'autoNumber', an object with a type of the same name, or undefined.`;
}
