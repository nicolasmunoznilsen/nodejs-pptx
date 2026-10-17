import { EMU_PER_INCH } from '../ooxml/constants.js';
import type { XmlNode } from '../ooxml/xml.js';
import { alignmentToOoxml, assertAlignment, type ParagraphAlignment } from '../text/alignment.js';
import { assertMarginInches, marginEmu, MARGIN_ATTR, MARGIN_OPTION, MARGIN_SIDES } from '../text/margins.js';

function inches(value: number): number { return Math.round(value * EMU_PER_INCH); }

export interface TextBoxOptions { x?: number; y?: number; width?: number; height?: number; fontSize?: number; color?: string; fontName?: string; name?: string; alignment?: ParagraphAlignment; margin?: number; marginLeft?: number; marginTop?: number; marginRight?: number; marginBottom?: number }
export interface TextBoxInfo { text: string; name: string; id: string; index: number; x: number; y: number; width: number; height: number; alignment?: ParagraphAlignment; margin?: number; marginLeft: number; marginTop: number; marginRight: number; marginBottom: number }

/** @internal The explicit `a:bodyPr` inset attributes for a set of options, `undefined` = inherit. */
export function bodyPrMarginAttrs(options: TextBoxOptions): XmlNode {
  const attrs: XmlNode = {};
  for (const side of MARGIN_SIDES) {
    const next = assertMarginInches(options[MARGIN_OPTION[side]] ?? options.margin);
    if (next !== undefined) attrs[`@_${MARGIN_ATTR[side]}`] = String(marginEmu(next));
  }
  return attrs;
}

export function textBoxXml(text: string, id: string, o: TextBoxOptions): XmlNode {
  const alignment = assertAlignment(o.alignment);
  const paragraphs = String(text).split('\n').map(line => {
    const isBullet = line.startsWith('• ');
    const textContent = isBullet ? line.substring(2) : line;
    const pPr: XmlNode = isBullet ? {
      '@_marL': '274320', '@_indent': '-274320',
      'a:buFont': { '@_typeface': o.fontName ?? 'Calibri' },
      'a:buChar': { '@_char': '•' },
    } : {};
    if (alignment) pPr['@_algn'] = alignmentToOoxml(alignment);
    return { 'a:pPr': pPr, 'a:r': { 'a:rPr': { '@_lang': 'en-US', '@_sz': String((o.fontSize ?? 18) * 100), 'a:solidFill': { 'a:srgbClr': { '@_val': (o.color ?? '000000').replace('#', '').toUpperCase() } }, 'a:latin': { '@_typeface': o.fontName ?? 'Calibri' } }, 'a:t': textContent } };
  });
  return { 'p:nvSpPr': { 'p:cNvPr': { '@_id': id, '@_name': o.name ?? `TextBox ${id}` }, 'p:cNvSpPr': { '@_txBox': '1' }, 'p:nvPr': {} }, 'p:spPr': { 'a:xfrm': { 'a:off': { '@_x': String(inches(o.x ?? 0.5)), '@_y': String(inches(o.y ?? 0.5)) }, 'a:ext': { '@_cx': String(inches(o.width ?? 4)), '@_cy': String(inches(o.height ?? 1)) } }, 'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} }, 'a:noFill': {}, 'a:ln': { 'a:noFill': {} } }, 'p:txBody': { 'a:bodyPr': { '@_wrap': 'square', ...bodyPrMarginAttrs(o) }, 'a:lstStyle': {}, 'a:p': paragraphs } };
}
