import { readAttr, writeAttr, type XmlNode } from '../ooxml/xml.js';
import { Shape } from './shape.js';
import { cellParagraphs, cellXml, emu, ensureFrameXfrm, firstObject, objects, tableNode, toEmu, toInches, type TableInfo } from './table-xml.js';

/** One `a:tc`: cell text (get/set) plus the raw XML node. */
export class TableCell {
  constructor(readonly node: XmlNode, private readonly markDirty: () => void) {}

  get text(): string { return this.paragraphNodes.map(node => runsText(node)).join('\n'); }
  /** Writes one `a:p` per line, reusing existing paragraphs (and their run props) where possible. */
  set text(value: string) {
    const lines = String(value).split('\n');
    const paragraphs = this.paragraphNodes;
    for (let index = 0; index < Math.min(lines.length, paragraphs.length); index += 1) {
      setParagraphText(paragraphs[index], lines[index]);
    }
    for (let index = paragraphs.length; index < lines.length; index += 1) {
      paragraphs.push(cellParagraphs(lines[index])[0]);
    }
    this.txBody['a:p'] = paragraphs.slice(0, lines.length);
    this.markDirty();
  }

  /** The `a:p` nodes of this cell. */
  get paragraphs(): XmlNode[] { return this.paragraphNodes; }
  get rawNode(): XmlNode { return this.node; }

  private get txBody(): XmlNode {
    const existing = firstObject(this.node['a:txBody']);
    if (existing) return existing;
    const created: XmlNode = { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': cellParagraphs('') };
    this.node['a:txBody'] = created;
    return created;
  }

  private get paragraphNodes(): XmlNode[] {
    const body = this.txBody;
    const paragraphs = objects(body['a:p']);
    body['a:p'] = paragraphs;
    return paragraphs;
  }
}

/** One `a:tr`: cell access, iteration, and row height in inches. */
export class TableRow {
  readonly cells: TableCell[];
  constructor(readonly node: XmlNode, private readonly markDirty: () => void, private readonly sync: () => void = () => {}) {
    this.cells = objects(node['a:tc']).map(cell => new TableCell(cell, markDirty));
  }
  get length(): number { return this.cells.length; }
  cell(index: number): TableCell {
    if (!Number.isInteger(index) || index < 0 || index >= this.cells.length) {
      throw new RangeError(`Table column index ${index} is out of range (0-${this.cells.length - 1}).`);
    }
    return this.cells[index];
  }
  get texts(): string[] { return this.cells.map(cell => cell.text); }
  [Symbol.iterator](): IterableIterator<TableCell> { return this.cells[Symbol.iterator](); }

  /** Row height (`a:tr/@h`), in inches. */
  get height(): number { return toInches(readAttr(this.node, 'h')); }
  set height(value: number) {
    writeAttr(this.node, 'h', toEmu(value, 'row height'));
    this.markDirty();
    this.sync();
  }
  get rawNode(): XmlNode { return this.node; }
}

/** One `a:gridCol`: column width in inches plus vertical cell access. */
export class TableColumn {
  constructor(readonly table: Table, readonly index: number) {}
  get node(): XmlNode { return this.table.gridColumnNodes[this.index]; }
  /** Column width (`a:gridCol/@w`), in inches. */
  get width(): number { return toInches(readAttr(this.node, 'w')); }
  set width(value: number) {
    writeAttr(this.node, 'w', toEmu(value, 'column width'));
    this.table.syncFrameExtent();
    this.table.onDirty();
  }
  get cells(): TableCell[] { return this.table.rowNodes.map(row => new TableRow(row, this.table.onDirty, () => this.table.syncFrameExtent()).cell(this.index)); }
  cell(rowIndex: number): TableCell { return this.table.cell(rowIndex, this.index); }
  get texts(): string[] { return this.cells.map(cell => cell.text); }
  [Symbol.iterator](): IterableIterator<TableCell> { return this.cells[Symbol.iterator](); }
  get length(): number { return this.cells.length; }
}

/**
 * Table facade over a `p:graphicFrame` holding `a:graphicData[@uri=...table]/a:tbl`.
 * Geometry (`x`, `y`, `width`, `height`) is reported in inches, unlike the raw EMU values
 * exposed by {@link Shape}.
 */
export class Table extends Shape {
  /** @internal Marks the owning slide dirty. */
  readonly onDirty: () => void;

  constructor(node: XmlNode, markDirty: () => void) {
    super(node, 'graphicFrame', markDirty);
    this.onDirty = markDirty;
  }

  /** @internal The `a:tbl` node inside this frame. */
  get tbl(): XmlNode {
    const tbl = tableNode(this.node);
    if (!tbl) throw new Error(`Shape ${this.id || this.name} does not contain a table.`);
    return tbl;
  }

  /** @internal The `a:tr` nodes, normalized to an array. */
  get rowNodes(): XmlNode[] { return objects(this.tbl['a:tr']); }
  /** @internal The `a:gridCol` nodes, normalized to an array. */
  get gridColumnNodes(): XmlNode[] { return objects(firstObject(this.tbl['a:tblGrid'])?.['a:gridCol']); }

  /** Row count; use {@link row}, {@link rowItems} or iteration to access rows. */
  get rows(): number { return this.rowNodes.length; }
  /** Column count; use {@link column}, {@link columnItems} or {@link forEachColumn}. */
  get columns(): number { return this.gridColumnNodes.length || objects(firstObject(this.rowNodes[0])?.['a:tc']).length; }

  get rowItems(): TableRow[] { return this.rowNodes.map(row => new TableRow(row, this.onDirty, () => this.syncFrameExtent())); }
  get columnItems(): TableColumn[] { return Array.from({ length: this.columns }, (_, index) => new TableColumn(this, index)); }

  row(index: number): TableRow {
    const rows = this.rowNodes;
    if (!Number.isInteger(index) || index < 0 || index >= rows.length) {
      throw new RangeError(`Table row index ${index} is out of range (0-${rows.length - 1}).`);
    }
    return new TableRow(rows[index], this.onDirty, () => this.syncFrameExtent());
  }
  column(index: number): TableColumn {
    const count = this.columns;
    if (!Number.isInteger(index) || index < 0 || index >= count) {
      throw new RangeError(`Table column index ${index} is out of range (0-${count - 1}).`);
    }
    return new TableColumn(this, index);
  }
  cell(rowIndex: number, columnIndex: number): TableCell {
    const row = this.row(rowIndex);
    if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= row.length) {
      throw new RangeError(`Table column index ${columnIndex} is out of range (0-${row.length - 1}).`);
    }
    return row.cell(columnIndex);
  }

  [Symbol.iterator](): IterableIterator<TableRow> { return this.rowItems[Symbol.iterator](); }
  forEachRow(callback: (row: TableRow, index: number) => void): void { this.rowItems.forEach(callback); }
  forEachColumn(callback: (column: TableColumn, index: number) => void): void { this.columnItems.forEach(callback); }

  /** All cell texts as a `string[][]`. */
  get data(): string[][] { return this.rowItems.map(row => row.texts); }
  set data(value: string[][]) {
    if (!Array.isArray(value) || value.length === 0) throw new RangeError('Table data must contain at least one row.');
    value.forEach((row, rowIndex) => {
      if (rowIndex >= this.rows) this.addRow(row);
      else row.forEach((text, columnIndex) => { this.cell(rowIndex, columnIndex).text = String(text ?? ''); });
    });
    this.onDirty();
  }

  /** Appends a row (reusing the last row height) and returns it. */
  addRow(values: string[] = []): TableRow {
    const rows = this.rowNodes;
    const width = this.columns;
    if (values.length > width) throw new RangeError(`Cannot add ${values.length} value(s) to a table with ${width} column(s).`);
    const height = readAttr(rows[rows.length - 1], 'h') ?? String(emu(1 / Math.max(1, rows.length)));
    const cells = Array.from({ length: width }, (_, column) => cellXml(values[column] ?? ''));
    this.tbl['a:tr'] = [...rows, { '@_h': height, 'a:tc': cells }];
    this.syncFrameExtent();
    this.onDirty();
    return new TableRow(this.rowNodes[this.rowNodes.length - 1], this.onDirty, () => this.syncFrameExtent());
  }

  /** Table position/size, in inches. */
  get x(): number { return toInches(readAttr(this.xfrmOff, 'x')); }
  set x(value: number) { writeAttr(this.xfrmOff, 'x', toEmu(value, 'x')); this.onDirty(); }
  get y(): number { return toInches(readAttr(this.xfrmOff, 'y')); }
  set y(value: number) { writeAttr(this.xfrmOff, 'y', toEmu(value, 'y')); this.onDirty(); }
  get width(): number { return toInches(readAttr(this.xfrmExt, 'cx')); }
  get height(): number { return toInches(readAttr(this.xfrmExt, 'cy')); }

  get info(): TableInfo {
    return {
      id: this.id,
      name: this.name,
      rows: this.rows,
      columns: this.columns,
      x: this.x,
      y: this.y,
      width: this.width,
      height: this.height,
      columnWidths: this.columnItems.map(column => column.width),
      rowHeights: this.rowItems.map(row => row.height),
      firstRow: this.tblFlag('firstRow'),
      banding: this.tblFlag('bandRow'),
      data: this.data,
    };
  }

  /** @internal Keeps the frame extent (`a:ext`) consistent with the grid and row sizes. */
  syncFrameExtent(): void {
    const cx = this.gridColumnNodes.reduce((total, node) => total + (Number(readAttr(node, 'w')) || 0), 0);
    const cy = this.rowNodes.reduce((total, node) => total + (Number(readAttr(node, 'h')) || 0), 0);
    if (cx) writeAttr(this.xfrmExt, 'cx', String(cx));
    if (cy) writeAttr(this.xfrmExt, 'cy', String(cy));
  }

  private tblFlag(name: string): boolean {
    const value = readAttr(firstObject(this.tbl['a:tblPr']), name);
    return value !== undefined && value !== '0';
  }
  /** The frame transform: existing `p:xfrm` (or legacy `a:xfrm`) is reused, never duplicated. */
  private get xfrmNode(): XmlNode { return ensureFrameXfrm(this.node); }
  private get xfrmOff(): XmlNode {
    const xfrm = this.xfrmNode;
    const off = firstObject(xfrm['a:off']);
    if (off) return off;
    const created: XmlNode = {};
    xfrm['a:off'] = created;
    return created;
  }
  private get xfrmExt(): XmlNode {
    const xfrm = this.xfrmNode;
    const ext = firstObject(xfrm['a:ext']);
    if (ext) return ext;
    const created: XmlNode = {};
    xfrm['a:ext'] = created;
    return created;
  }
}

function runsText(paragraph: XmlNode): string {
  return objects(paragraph['a:r']).map(run => {
    const text = run['a:t'];
    return typeof text === 'string' || typeof text === 'number' ? String(text) : '';
  }).join('');
}

function setParagraphText(paragraph: XmlNode, value: string): void {
  const runs = objects(paragraph['a:r']);
  if (runs.length) {
    runs[0]['a:t'] = value;
    for (const run of runs.slice(1)) run['a:t'] = '';
    paragraph['a:r'] = runs.length === 1 ? runs[0] : runs;
    return;
  }
  const rest = Object.fromEntries(Object.entries(paragraph).filter(([key]) => key !== 'a:r'));
  for (const key of Object.keys(paragraph)) delete paragraph[key];
  for (const [key, entry] of Object.entries(rest)) paragraph[key] = entry;
  paragraph['a:r'] = { 'a:rPr': { '@_lang': 'en-US' }, 'a:t': value };
}
