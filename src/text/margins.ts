import { EMU_PER_INCH } from '../ooxml/constants.js';
import { readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';

/**
 * The four text-frame insets of `a:bodyPr`, mirroring the `lIns`/`tIns`/`rIns`/`bIns` attributes
 * and the `MSO_TEXT_SIDE`-style members used by `python-pptx`.
 */
export type TextFrameMarginSide = 'left' | 'top' | 'right' | 'bottom';

/** Public side -> OOXML attribute name. Frozen so the lookup tables cannot be mutated at runtime. */
export const MARGIN_ATTR: Readonly<Record<TextFrameMarginSide, string>> = Object.freeze({
  left: 'lIns',
  top: 'tIns',
  right: 'rIns',
  bottom: 'bIns',
});

/**
 * Public side -> inset PowerPoint applies when the attribute is absent, in EMU. These are the
 * DrawingML defaults (`lIns`/`rIns` = 91440 = 0.1", `tIns`/`bIns` = 45720 = 0.05").
 */
export const MARGIN_DEFAULT_EMU: Readonly<Record<TextFrameMarginSide, number>> = Object.freeze({
  left: 91440,
  top: 45720,
  right: 91440,
  bottom: 45720,
});

/** Public side -> `TextBoxOptions` key. Frozen so the lookup tables cannot be mutated at runtime. */
export const MARGIN_OPTION: Readonly<Record<TextFrameMarginSide, 'marginLeft' | 'marginTop' | 'marginRight' | 'marginBottom'>> = Object.freeze({
  left: 'marginLeft',
  top: 'marginTop',
  right: 'marginRight',
  bottom: 'marginBottom',
});

/** Every supported side, in OOXML attribute order. */
export const MARGIN_SIDES: readonly TextFrameMarginSide[] = Object.freeze(['left', 'top', 'right', 'bottom']);

/**
 * Public side -> OOXML attribute name, resolved through an own-property lookup.
 *
 * The guard is `Object.hasOwn`, not `in` and not a truthiness check, so names inherited from
 * `Object.prototype` (`constructor`, `toString`, `valueOf`, `__proto__`, ...) never resolve
 * through the prototype chain and never leak an inherited value into generated XML.
 */
export function marginAttrFor(side: unknown): string | undefined {
  if (typeof side !== 'string' || !Object.hasOwn(MARGIN_ATTR, side)) return undefined;
  return (MARGIN_ATTR as Record<string, string>)[side];
}

/** @internal Default inset for a side, in inches, or `undefined` for an unknown side. */
function defaultMarginInches(side: unknown): number | undefined {
  if (typeof side !== 'string' || !Object.hasOwn(MARGIN_DEFAULT_EMU, side)) return undefined;
  return (MARGIN_DEFAULT_EMU as Record<string, number>)[side] / EMU_PER_INCH;
}

/** @internal Turn an OOXML EMU attribute into a finite, non-negative EMU count, if it is one. */
function explicitEmu(bodyPr: XmlNode | undefined, attr: string): number | undefined {
  const raw = readAttr(bodyPr, attr);
  if (typeof raw !== 'string' || raw.trim() === '') return undefined;
  const emu = Number(raw);
  return Number.isFinite(emu) && emu >= 0 ? emu : undefined;
}

/**
 * The explicit `@_*Ins` value of `bodyPr`, in EMU, or `undefined` when the attribute is missing
 * or carries a value this library does not understand (which means "inherit the default").
 *
 * Used by the public setters to decide whether an assignment actually changes anything, so a
 * no-op write never marks the slide dirty.
 */
export function explicitMarginEmu(bodyPr: XmlNode | undefined, side: TextFrameMarginSide): number | undefined {
  const attr = marginAttrFor(side);
  if (attr === undefined) throw new TypeError(describeInvalidSide(side));
  return explicitEmu(bodyPr, attr);
}

/**
 * Reads one text-frame inset of an `a:bodyPr` node, in inches. An absent or malformed attribute
 * reports the OOXML default for that side (0.1" left/right, 0.05" top/bottom), which is exactly
 * what PowerPoint renders, so the result is always a usable number.
 */
export function readMarginInches(bodyPr: XmlNode | undefined, side: TextFrameMarginSide): number {
  const attr = marginAttrFor(side);
  if (attr === undefined) throw new TypeError(describeInvalidSide(side));
  const emu = explicitEmu(bodyPr, attr);
  if (emu !== undefined) return emu / EMU_PER_INCH;
  return defaultMarginInches(side) as number;
}

/**
 * Writes or clears (`undefined`) one text-frame inset of an `a:bodyPr` node. The value is stored
 * as a rounded EMU string; `undefined` removes the attribute so the side inherits the PowerPoint
 * default again.
 */
export function writeMarginInches(
  bodyPr: XmlNode,
  side: TextFrameMarginSide,
  inches: number | undefined,
): void {
  const attr = marginAttrFor(side);
  if (attr === undefined) throw new TypeError(describeInvalidSide(side));
  const value = assertMarginInches(inches);
  if (value === undefined) {
    writeAttr(bodyPr, attr, undefined);
    return;
  }
  writeAttr(bodyPr, attr, String(marginEmu(value)));
}

/** @internal Inches -> EMU, rounded to the nearest whole EMU that OOXML can store. */
export function marginEmu(inches: number): number {
  return Math.round(inches * EMU_PER_INCH);
}

/** @internal Shared message for rejected margin sides. */
function describeInvalidSide(side: unknown): string {
  return `Invalid text-frame margin side: ${JSON.stringify(side)}. Expected one of: ${Object.keys(MARGIN_ATTR).join(', ')}.`;
}

/** @internal Shared message for rejected margin values. */
function describeInvalidInches(value: unknown): string {
  return `Invalid text-frame margin: ${JSON.stringify(value)}. Expected a finite, non-negative number of inches, or undefined to inherit the default.`;
}

/** Throws a descriptive error when `value` is not a supported side. */
export function assertMarginSide(value: unknown): TextFrameMarginSide {
  if (marginAttrFor(value) === undefined) throw new TypeError(describeInvalidSide(value));
  return value as TextFrameMarginSide;
}

/**
 * Throws a descriptive error when `value` is not a usable inset. `undefined` and `null` are
 * accepted and mean "inherit the default".
 */
export function assertMarginInches(value: unknown): number | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) {
    throw new TypeError(describeInvalidInches(value));
  }
  return value;
}
