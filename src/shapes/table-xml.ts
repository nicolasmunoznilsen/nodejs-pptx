import { EMU_PER_INCH } from '../ooxml/constants.js';
import { arr, readAttr, type XmlNode } from '../ooxml/xml.js';

/** `a:graphicData/@uri` value that identifies a DrawingML table inside a `p:graphicFrame`. */
export const TABLE_GRAPHIC_DATA_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

export interface TableOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  columnWidths?: number[];
  rowHeights?: number[];
  name?: string;
  firstRow?: boolean;
  banding?: boolean;
}

export interface TableInfo {
  id: string;
  name: string;
  rows: number;
  columns: number;
  x: number;
  y: number;
  width: number;
  height: number;
  columnWidths: number[];
  rowHeights: number[];
  firstRow: boolean;
  banding: boolean;
  data: string[][];
}

/** Inches to EMU, rounded to the nearest integer. */
export function emu(value: number): number { return Math.round(value * EMU_PER_INCH); }
/** Validates an inch value and returns it as an EMU string. */
export function toEmu(value: number, label: string): string {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    throw new RangeError(`Table ${label} must be a positive finite number of inches.`);
  }
  return String(emu(value));
}
/** EMU string (or number) to inches. */
export function toInches(raw: string | number | undefined): number { return (Number(raw ?? 0) || 0) / EMU_PER_INCH; }
function sum(values: number[]): number { return values.reduce((total, value) => total + value, 0); }

/** Normalizes user input into a rectangular matrix of strings. */
export function normalizeTableData(data: string[][]): string[][] {
  if (!Array.isArray(data) || data.length === 0) throw new RangeError('Table data must contain at least one row.');
  const width = Math.max(...data.map(row => {
    if (!Array.isArray(row)) throw new RangeError('Every table row must be an array of strings.');
    return row.length;
  }));
  if (width === 0) throw new RangeError('Table data must contain at least one column.');
  return data.map(row => Array.from({ length: width }, (_, column) => row[column] === undefined || row[column] === null ? '' : String(row[column])));
}

/** Splits `total` (EMU) into `count` slices, or validates the explicit EMU list provided. */
function slices(values: number[] | undefined, count: number, total: number, label: string): number[] {
  if (values === undefined) {
    const base = Math.max(1, Math.floor(total / count));
    const out = Array.from({ length: count }, () => base);
    out[count - 1] = Math.max(1, total - base * (count - 1));
    return out;
  }
  if (!Array.isArray(values) || values.length !== count) {
    throw new RangeError(`Table ${label} must contain exactly ${count} value(s), got ${Array.isArray(values) ? values.length : 0}.`);
  }
  return values.map(value => {
    if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
      throw new RangeError(`Table ${label} values must be positive finite numbers of inches.`);
    }
    return emu(value);
  });
}

/** One `a:p` per line, each holding a single run so `cell.text` round-trips. */
export function cellParagraphs(text: string): XmlNode[] {
  return String(text).split('\n').map(line => ({ 'a:r': { 'a:rPr': { '@_lang': 'en-US' }, 'a:t': line } }));
}

/** One `a:tc` holding a `txBody` with one `a:p` per line. */
export function cellXml(text: string): XmlNode {
  return { 'a:txBody': { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': cellParagraphs(text) } };
}

/** Builds a `p:graphicFrame` wrapping a DrawingML table, ready to append to `p:spTree`. */
export function tableXml(data: string[][], id: string, o: TableOptions = {}): XmlNode {
  const rows = normalizeTableData(data);
  if (rows.length === 0) throw new RangeError('Table data must contain at least one row.');
  const columnCount = rows[0].length;
  const width = o.width ?? 6;
  const height = o.height ?? 1;
  if (typeof width !== 'number' || !Number.isFinite(width) || width <= 0) throw new RangeError('Table width must be a positive finite number of inches.');
  if (typeof height !== 'number' || !Number.isFinite(height) || height <= 0) throw new RangeError('Table height must be a positive finite number of inches.');
  const columnWidths = slices(o.columnWidths, columnCount, emu(width), 'columnWidths');
  const rowHeights = slices(o.rowHeights, rows.length, emu(height), 'rowHeights');
  return {
    'p:nvGraphicFramePr': {
      'p:cNvPr': { '@_id': id, '@_name': o.name ?? `Table ${id}` },
      'p:nvGraphicFramePr': {},
      'p:nvPr': {},
    },
    // CT_GraphicFrame requires the presentationml-namespace transform (`p:xfrm`),
    // positioned between p:nvGraphicFramePr and a:graphic.
    'p:xfrm': {
      'a:off': { '@_x': String(emu(o.x ?? 1)), '@_y': String(emu(o.y ?? 1)) },
      'a:ext': { '@_cx': String(sum(columnWidths)), '@_cy': String(sum(rowHeights)) },
    },
    'a:graphic': {
      'a:graphicData': {
        '@_uri': TABLE_GRAPHIC_DATA_URI,
        'a:tbl': {
          'a:tblPr': { '@_firstRow': o.firstRow === false ? '0' : '1', '@_bandRow': o.banding ? '1' : '0' },
          'a:tblGrid': { 'a:gridCol': columnWidths.map(value => ({ '@_w': String(value) })) },
          'a:tr': rows.map((row, index) => ({ '@_h': String(rowHeights[index]), 'a:tc': row.map(cellXml) })),
        },
      },
    },
  };
}

/** True when a `p:graphicFrame` node hosts a DrawingML table. */
export function isTableGraphicFrame(frame: XmlNode): boolean {
  return tableNode(frame) !== undefined;
}

/** The `a:tbl` node of a `p:graphicFrame`, or `undefined` when it is not a table. */
export function tableNode(frame: unknown): XmlNode | undefined {
  if (!frame || typeof frame !== 'object') return undefined;
  const node = frame as XmlNode;
  for (const graphic of objects(node['a:graphic'])) {
    for (const data of objects(graphic['a:graphicData'])) {
      if (readAttr(data, 'uri') !== TABLE_GRAPHIC_DATA_URI) continue;
      const tbl = data['a:tbl'];
      const first = Array.isArray(tbl) ? tbl[0] : tbl;
      if (first && typeof first === 'object') return first as XmlNode;
    }
  }
  return undefined;
}

/** Keeps only object entries: parsed XML may hold strings or arrays depending on cardinality. */
export function objects(value: unknown): XmlNode[] {
  return arr(value as XmlNode | XmlNode[] | undefined).filter((item): item is XmlNode => Boolean(item) && typeof item === 'object');
}

/** The first object entry of a parsed child, or `undefined` (a single node parses to an object, not an array). */
export function firstObject(value: unknown): XmlNode | undefined {
  const first = Array.isArray(value) ? value[0] : value;
  return first && typeof first === 'object' ? (first as XmlNode) : undefined;
}

/**
 * The transform of a `p:graphicFrame`: the presentationml `p:xfrm` that PowerPoint/LibreOffice emit,
 * falling back to a legacy `a:xfrm` only when that is what the part actually holds.
 */
export function frameXfrm(frame: XmlNode): XmlNode | undefined {
  return firstObject(frame['p:xfrm']) ?? firstObject(frame['a:xfrm']);
}

/**
 * The transform of a `p:graphicFrame`, created when missing. An existing node is always returned
 * as-is (and mutated in place) so open→edit→save never yields two transforms in one frame; a newly
 * created `p:xfrm` is inserted after `p:nvGraphicFramePr` to keep CT_GraphicFrame order valid.
 */
export function ensureFrameXfrm(frame: XmlNode): XmlNode {
  const existing = frameXfrm(frame);
  if (existing) return existing;
  const created: XmlNode = {};
  // Drop non-element placeholders first: a self-closed `<p:xfrm/>` parses to "" and would
  // otherwise clobber the node created below and yield two transforms when serialized.
  const keys: string[] = [];
  for (const key of Object.keys(frame)) {
    if (key === 'p:xfrm' || key === 'a:xfrm') { delete frame[key]; continue; }
    keys.push(key);
  }
  const anchor = keys.indexOf('p:nvGraphicFramePr');
  const ordered = new Map<string, unknown>(keys.map(key => [key, frame[key]]));
  for (const key of keys) delete frame[key];
  let inserted = false;
  const insert = (): void => {
    if (inserted) return;
    frame['p:xfrm'] = created;
    inserted = true;
  };
  for (const [key, value] of ordered) {
    // Insert right after p:nvGraphicFramePr, or before a:graphic when the frame lacks that part.
    if (anchor === -1 ? key === 'a:graphic' : key === 'p:nvGraphicFramePr') insert();
    frame[key] = value;
  }
  insert();
  return created;
}
