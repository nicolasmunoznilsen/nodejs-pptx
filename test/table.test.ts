import JSZip from 'jszip';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Presentation, Table, TableCell, TableColumn, TableRow } from '../src/index.js';

describe('Slide.addTable', () => {
  it('creates a table with DrawingML order and default geometry', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const table = slide.addTable([['A', 'B'], ['C', 'D']], { x: 1, y: 2, width: 6, height: 1.5, name: 'Grid' });
    expect(table).toBeInstanceOf(Table);
    expect(table.rows).toBe(2);
    expect(table.columns).toBe(2);
    expect(table.name).toBe('Grid');
    expect(table.x).toBeCloseTo(1);
    expect(table.y).toBeCloseTo(2);
    expect(table.width).toBeCloseTo(6);
    expect(table.height).toBeCloseTo(1.5);
    expect(table.data).toEqual([['A', 'B'], ['C', 'D']]);
    expect(table.cell(0, 1).text).toBe('B');

    const xml = await (await JSZip.loadAsync(await pptx.toBuffer())).file('ppt/slides/slide1.xml')!.async('string');
    const frame = /<p:graphicFrame>.*<\/p:graphicFrame>/.exec(xml)![0];
    const order = ['p:nvGraphicFramePr', 'p:cNvPr', 'p:nvGraphicFramePr', 'p:nvPr', 'p:xfrm', 'a:off', 'a:ext', 'a:graphic', 'a:graphicData', 'a:tbl', 'a:tblPr', 'a:tblGrid', 'a:gridCol', 'a:tr', 'a:tc', 'a:txBody', 'a:bodyPr', 'a:lstStyle', 'a:p', 'a:r', 'a:t'];
    let cursor = -1;
    for (const tag of order) {
      const found = frame.indexOf(`<${tag}`, cursor + 1);
      expect(found, `${tag} after previous sibling in ${frame}`).toBeGreaterThan(cursor);
      cursor = found;
    }
    expect(frame).toContain('a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"');
    expect(frame.match(/<a:gridCol/g)).toHaveLength(2);
    expect(frame.match(/<a:tr/g)).toHaveLength(2);
    expect(frame).not.toContain('Relationship');
  });

  it('supports explicit column widths and row heights in inches', () => {
    const slide = Presentation.create().slides.add();
    const table = slide.addTable([['a', 'b'], ['c', 'd']], { width: 6, height: 2, columnWidths: [2, 4], rowHeights: [0.5, 1.5] });
    expect(table.columnItems.map(c => c.width)).toEqual([2, 4]);
    expect(table.rowItems.map(r => r.height)).toEqual([0.5, 1.5]);
    expect(table.width).toBeCloseTo(6);
    expect(table.height).toBeCloseTo(2);
    expect(() => slide.addTable([['a', 'b']], { columnWidths: [1] })).toThrow(/columnWidths/);
    table.column(0).width = 3;
    expect(table.column(0).width).toBeCloseTo(3);
    expect(table.width).toBeCloseTo(7);
    table.row(0).height = 1;
    expect(table.row(0).height).toBeCloseTo(1);
    expect(table.height).toBeCloseTo(2.5);
  });

  it('round-trips 1x1 tables and single-cell rows through the zip', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTable([['solo']]);
    const reopened = await Presentation.open(await pptx.toBuffer());
    const table = reopened.slides.get(0).tables[0];
    expect(table).toBeInstanceOf(Table);
    expect(table.rows).toBe(1);
    expect(table.columns).toBe(1);
    expect(table.cell(0, 0).text).toBe('solo');
    expect(table.data).toEqual([['solo']]);
    expect(table.rowItems[0]).toBeInstanceOf(TableRow);
    expect(table.cell(0, 0)).toBeInstanceOf(TableCell);
    expect(table.column(0)).toBeInstanceOf(TableColumn);
    table.cell(0, 0).text = 'uno\ndos\ntres';
    expect(table.cell(0, 0).text).toBe('uno\ndos\ntres');
    const again = await Presentation.open(await reopened.toBuffer());
    expect(again.slides.get(0).tables[0].data).toEqual([['uno\ndos\ntres']]);
    expect(again.slides.get(0).tables[0].rows).toBe(1);
  });

  it('edits text, adds rows, and preserves run properties', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const table = slide.addTable([['Head 1', 'Head 2'], ['a', 'b']], { banding: true, firstRow: true });
    const cell = table.cell(0, 0);
    cell.text = 'Renamed';
    expect(cell.text).toBe('Renamed');
    const added = table.addRow(['c', 'd']);
    expect(table.rows).toBe(3);
    expect(added.texts).toEqual(['c', 'd']);
    expect(table.data).toEqual([['Renamed', 'Head 2'], ['a', 'b'], ['c', 'd']]);
    const reopened = await Presentation.open(await pptx.toBuffer());
    const fresh = reopened.slides.get(0).tables[0];
    expect(fresh.data).toEqual([['Renamed', 'Head 2'], ['a', 'b'], ['c', 'd']]);
    const run = fresh.cell(1, 0).paragraphs[0]['a:r'] as Record<string, unknown>;
    (run as { 'a:rPr': Record<string, unknown> })['a:rPr'] = { '@_sz': '2400' };
    fresh.cell(1, 0).text = 'styled';
    expect((fresh.cell(1, 0).paragraphs[0]['a:r'] as { 'a:rPr': Record<string, unknown> })['a:rPr']['@_sz']).toBe('2400');
    expect(fresh.cell(1, 0).text).toBe('styled');
    expect(fresh.info.banding).toBe(true);
    expect(fresh.info.firstRow).toBe(true);
  });

  it('throws RangeError for out-of-range row, column, and cell access', () => {
    const slide = Presentation.create().slides.add();
    const table = slide.addTable([['a', 'b']]);
    expect(() => table.cell(1, 0)).toThrow(/row index 1 is out of range \(0-0\)/);
    expect(() => table.cell(0, 2)).toThrow(/column index 2 is out of range \(0-1\)/);
    expect(() => table.row(-1)).toThrow(/out of range/);
    expect(() => table.column(5)).toThrow(/out of range/);
    expect(() => table.row(0).cell(9)).toThrow(/out of range/);
    expect(() => table.addRow(['a', 'b', 'c'])).toThrow(/Cannot add 3 value/);
    expect(() => slide.addTable([])).toThrow(/at least one row/);
  });

  it('exposes tables through the shape collection and slide helpers', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('title');
    slide.addTable([['1', '2'], ['3', '4']], { name: 'Data' });
    expect(slide.shapes.length).toBe(2);
    const frame = slide.shapes.findByName('Data');
    expect(frame).toBeInstanceOf(Table);
    expect(slide.shapes.tables).toHaveLength(1);
    expect([...slide.shapes].some(shape => shape instanceof Table)).toBe(true);
    expect(slide.getTables()[0]).toMatchObject({ name: 'Data', rows: 2, columns: 2, data: [['1', '2'], ['3', '4']] });
    const rows: string[] = [];
    for (const row of slide.tables[0]) rows.push(row.texts.join('|'));
    expect(rows).toEqual(['1|2', '3|4']);
    const columns: number[] = [];
    slide.tables[0].forEachColumn(column => columns.push(column.cells.length));
    expect(columns).toEqual([2, 2]);
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.get(0).text).toBe('title');
    expect(reopened.slides.get(0).tables[0].data).toEqual([['1', '2'], ['3', '4']]);
  });

  it('keeps non-table graphic frames as plain shapes', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    xml = xml.replace('</p:spTree>', '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="43" name="Chart 43"/><p:nvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame></p:spTree>');
    zip.file('ppt/slides/slide1.xml', xml);
    const opened = await Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
    const shape = opened.slides.get(0).shapes.findByName('Chart 43')!;
    expect(shape).not.toBeInstanceOf(Table);
    expect(shape.type).toBe('graphicFrame');
    expect(opened.slides.get(0).tables).toHaveLength(0);
    opened.slides.get(0).addTable([['after']]);
    expect(opened.slides.get(0).shapes.length).toBe(2);
    const saved = await JSZip.loadAsync(await opened.toBuffer());
    const savedXml = await saved.file('ppt/slides/slide1.xml')!.async('string');
    expect(savedXml).toContain('Chart 43');
    expect(savedXml.match(/<p:graphicFrame/g)).toHaveLength(2);
  });
});

// Independent acceptance verification for issue #22 (Tester). Assertions are written against
// the documented behavior in README.md; nothing here weakens the tests above.
describe('issue #22 acceptance: basic PPTX tables', () => {
  const EMU = 914400;

  async function slideXmlOf(pptx: Presentation): Promise<string> {
    return (await JSZip.loadAsync(await pptx.toBuffer())).file('ppt/slides/slide1.xml')!.async('string');
  }

  function tblOf(xml: string): string {
    const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(xml);
    expect(frame, 'slide XML must contain a p:graphicFrame').not.toBeNull();
    const tbl = /<a:tbl>[\s\S]*<\/a:tbl>/.exec(frame![0]);
    expect(tbl, 'the graphicFrame must contain a:tbl').not.toBeNull();
    return tbl![0];
  }

  it('(1) adds a 3x2 table whose a:tbl nests a:tblPr, a:tblGrid, then a:tr', async () => {
    const pptx = Presentation.create();
    const table = pptx.slides.add().addTable([['A', 'B'], ['C', 'D'], ['E', 'F']]);
    expect(table.rows).toBe(3);
    expect(table.columns).toBe(2);

    const tbl = tblOf(await slideXmlOf(pptx));
    const tblPr = tbl.indexOf('<a:tblPr');
    const tblGrid = tbl.indexOf('<a:tblGrid');
    const firstRow = tbl.indexOf('<a:tr');
    expect(tblPr).toBeGreaterThan(-1);
    expect(tblGrid).toBeGreaterThan(tblPr);
    expect(firstRow).toBeGreaterThan(tblGrid);
    expect(tbl.match(/<a:gridCol/g)).toHaveLength(2);
    expect(tbl.match(/<a:tr/g)).toHaveLength(3);
  });

  it('(2) re-reads every cell text and the row/column counts after toBuffer()', async () => {
    const data = [['A', 'B'], ['C', 'D'], ['E', 'F']];
    const pptx = Presentation.create();
    pptx.slides.add().addTable(data, { name: 'Grid' });
    const reopened = await Presentation.open(await pptx.toBuffer());
    const table = reopened.slides.get(0).tables[0];
    expect(table.rows).toBe(3);
    expect(table.columns).toBe(2);
    expect(table.data).toEqual(data);
    for (const [row, cells] of data.entries()) {
      for (const [column, text] of cells.entries()) expect(table.cell(row, column).text).toBe(text);
    }
  });

  it('(3) round-trips edited and multi-line cell text and keeps sibling textbox formatting', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Keep me', { name: 'Sibling', fontSize: 24, color: 'FF0000', fontName: 'Arial', alignment: 'center' });
    const table = slide.addTable([['old', 'b'], ['c', 'd']], { name: 'Grid' });
    table.cell(0, 0).text = 'new';
    table.cell(1, 1).text = 'x\ny';

    const reopened = await Presentation.open(await pptx.toBuffer());
    const fresh = reopened.slides.get(0).tables[0];
    expect(fresh.cell(0, 0).text).toBe('new');
    expect(fresh.cell(1, 1).text).toBe('x\ny');
    expect(fresh.data).toEqual([['new', 'b'], ['c', 'x\ny']]);

    const sibling = reopened.slides.get(0).shapes.findByName('Sibling')!;
    expect(sibling.textFrame.text).toBe('Keep me');
    const run = sibling.textFrame.paragraphs[0].runs[0];
    expect(run.font.size).toBe(24);
    expect(run.font.color).toBe('FF0000');
    expect(run.font.name).toBe('Arial');
    expect(sibling.textFrame.paragraphs[0].alignment).toBe('center');
    expect(await slideXmlOf(reopened)).toContain('<a:ext cx="5486400"');
  });

  it('(4) writes the expected gridCol/tr EMU values and keeps the grid sum equal to ext cx', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    const table = slide.addTable([['a', 'b'], ['c', 'd']], { width: 6, height: 2, columnWidths: [2, 4], rowHeights: [0.5, 1.5] });
    table.addRow(['e', 'f']);
    expect(table.rows).toBe(3);

    const xml = await slideXmlOf(pptx);
    const tbl = tblOf(xml);
    const widths = [...tbl.matchAll(/<a:gridCol[^>]*\bw="(\d+)"/g)].map(match => Number(match[1]));
    const heights = [...tbl.matchAll(/<a:tr[^>]*\bh="(\d+)"/g)].map(match => Number(match[1]));
    expect(widths).toEqual([2 * EMU, 4 * EMU]);
    expect(heights).toEqual([0.5 * EMU, 1.5 * EMU, 1.5 * EMU]);
    const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(xml)![0];
    const ext = /<a:ext[^>]*\bcx="(\d+)"/.exec(frame)!;
    expect(widths.reduce((total, value) => total + value, 0)).toBe(Number(ext[1]));
    expect(table.width).toBeCloseTo(6);
    expect(table.columnItems.map(column => column.width)).toEqual([2, 4]);
    expect(table.rowItems.map(row => row.height)).toEqual([0.5, 1.5, 1.5]);

    // Default geometry must also keep the grid and the frame extent in sync.
    const plain = Presentation.create();
    plain.slides.add().addTable([['1', '2', '3']], { width: 6, height: 1 });
    const plainXml = await slideXmlOf(plain);
    const plainTbl = tblOf(plainXml);
    const plainWidths = [...plainTbl.matchAll(/<a:gridCol[^>]*\bw="(\d+)"/g)].map(match => Number(match[1]));
    expect(plainWidths).toHaveLength(3);
    expect(plainWidths.reduce((total, value) => total + value, 0)).toBe(6 * EMU);
    expect(new RegExp(`<a:ext[^>]*\\bcx="${6 * EMU}"`).exec(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(plainXml)![0])).not.toBeNull();
  });

  it('(5) round-trips 1x1, single-column, and single-row tables', async () => {
    for (const data of [[['solo']], [['a'], ['b'], ['c']], [['only one row', 'second']]] as Array<string[][]>) {
      const pptx = Presentation.create();
      pptx.slides.add().addTable(data);
      const reopened = await Presentation.open(await pptx.toBuffer());
      const table = reopened.slides.get(0).tables[0];
      expect(table.data).toEqual(data);
      expect(table.rows).toBe(data.length);
      expect(table.columns).toBe(data[0].length);
      expect(table.cell(0, 0).text).toBe(data[0][0]);
      expect(table.rowItems.every(row => row.cells.length === data[0].length)).toBe(true);
    }
  });

  it('(6) throws RangeError for out-of-range row, column, and cell access', () => {
    const table = Presentation.create().slides.add().addTable([['a', 'b'], ['c', 'd']]);
    expect(() => table.cell(2, 0)).toThrow(RangeError);
    expect(() => table.cell(0, 2)).toThrow(RangeError);
    expect(() => table.row(2)).toThrow(RangeError);
    expect(() => table.column(2)).toThrow(RangeError);
    expect(() => table.row(0).cell(2)).toThrow(RangeError);
    expect(() => table.column(0).cell(2)).toThrow(RangeError);
  });

  it('(7) exposes typed tables in slide.shapes and leaves the textbox APIs unchanged', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Hello world', { name: 'Greeting' });
    slide.addTable([['A', 'B'], ['C', 'D']], { name: 'Grid' });

    expect(slide.shapes.length).toBe(2);
    const shape = slide.shapes.get(1);
    expect(shape).toBeInstanceOf(Table);
    expect(shape.type).toBe('graphicFrame');
    expect((shape as Table).cell(1, 1).text).toBe('D');
    expect(slide.shapes.tables).toHaveLength(1);
    expect(slide.getTables()[0]).toMatchObject({ name: 'Grid', rows: 2, columns: 2 });

    expect(slide.getTextBoxes()).toHaveLength(1);
    expect(slide.getTextBoxes()[0]).toMatchObject({ name: 'Greeting', text: 'Hello world', index: 0 });
    expect(slide.replaceText('world', 'PPTX')).toBe(1);
    expect(slide.getTextBoxes()[0].text).toBe('Hello PPTX');
    expect(slide.tables[0].data).toEqual([['A', 'B'], ['C', 'D']]);

    const reopened = await Presentation.open(await pptx.toBuffer());
    const reopenedSlide = reopened.slides.get(0);
    expect(reopenedSlide.getTextBoxes()).toHaveLength(1);
    expect(reopenedSlide.getTextBoxes()[0]).toMatchObject({ name: 'Greeting', text: 'Hello PPTX' });
    expect(reopenedSlide.shapes.get(1)).toBeInstanceOf(Table);
    expect(reopenedSlide.replaceText('A', 'Z')).toBe(0);
    expect(reopenedSlide.tables[0].data).toEqual([['A', 'B'], ['C', 'D']]);
  });

  it('(8) does not duplicate table text into slide.text', async () => {
    const pptx = Presentation.create();
    const slide = pptx.slides.add();
    slide.addTextBox('Narration', { name: 'Notes' });
    slide.addTable([['A', 'B'], ['C', 'D']]);
    expect(slide.text).toBe('Narration');
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.get(0).text).toBe('Narration');
    expect(reopened.slides.get(0).text).not.toContain('A');
    expect(reopened.slides.get(0).getTextBoxes().map(box => box.text)).toEqual(['Narration']);
  });

  it('round-trips tables through save()/open() on disk', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'nodejs-pptx-table-'));
    const file = join(dir, 'tables.pptx');
    const pptx = Presentation.create();
    pptx.slides.add().addTable([['x', 'y'], ['p', 'q']], { name: 'Disk' });
    await pptx.save(file);
    expect((await readFile(file)).length).toBeGreaterThan(0);
    const reopened = await Presentation.open(file);
    expect(reopened.slides.get(0).tables[0].data).toEqual([['x', 'y'], ['p', 'q']]);
  });
});

// CT_GraphicFrame requires the presentationml-namespace transform (`p:xfrm`), which is what
// PowerPoint and LibreOffice write. Reading and editing such a part must never produce a second
// transform element inside the same graphicFrame.
describe('issue #22: graphicFrame transform namespace', () => {
  const EMU = 914400;

  function tableFrameXml(xfrmTag: string, x: number, y: number, cx: number, cy: number): string {
    return `<${xfrmTag}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></${xfrmTag}>`
      + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
      + '<a:tbl><a:tblPr firstRow="1" bandRow="0"/><a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid>'
      + '<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>a</a:t></a:r></a:p></a:txBody></a:tc>'
      + '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>b</a:t></a:r></a:p></a:txBody></a:tc></a:tr>'
      + '<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>c</a:t></a:r></a:p></a:txBody></a:tc>'
      + '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>d</a:t></a:r></a:p></a:txBody></a:tc></a:tr>'
      + '</a:tbl></a:graphicData></a:graphic>';
  }

  async function slideWithFrame(frameInner: string): Promise<Presentation> {
    const pptx = Presentation.create();
    pptx.slides.add();
    const zip = await JSZip.loadAsync(await pptx.toBuffer());
    const path = 'ppt/slides/slide1.xml';
    let xml = await zip.file(path)!.async('string');
    xml = xml.replace('</p:spTree>', `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Real"/><p:nvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${frameInner}</p:graphicFrame></p:spTree>`);
    zip.file(path, xml);
    return Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
  }

  async function frameOf(pptx: Presentation): Promise<string> {
    const xml = await (await JSZip.loadAsync(await pptx.toBuffer())).file('ppt/slides/slide1.xml')!.async('string');
    return /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(xml)![0];
  }

  /** A CT_GraphicFrame holds exactly one transform, positioned before a:graphic. */
  function expectSingleTransform(frame: string): void {
    const transforms = [...frame.matchAll(/<(p:xfrm|a:xfrm)\b/g)];
    expect(transforms, frame).toHaveLength(1);
    const graphic = frame.indexOf('<a:graphic');
    expect(graphic).toBeGreaterThan(-1);
    expect(transforms[0].index).toBeLessThan(graphic);
  }

  it('reads geometry from a presentationml p:xfrm written by PowerPoint/LibreOffice', async () => {
    const opened = await slideWithFrame(tableFrameXml('p:xfrm', 2 * EMU, 3 * EMU, 5486400, 1828800));
    const table = opened.slides.get(0).tables[0];
    expect(table).toBeInstanceOf(Table);
    expect(table.name).toBe('Real');
    expect(table.x).toBeCloseTo(2);
    expect(table.y).toBeCloseTo(3);
    expect(table.width).toBeCloseTo(6);
    expect(table.height).toBeCloseTo(2);
    expect(table.x * EMU).toBe(2 * EMU);
    expect(table.y * EMU).toBe(3 * EMU);
    expect(table.data).toEqual([['a', 'b'], ['c', 'd']]);
  });

  it('still reads a legacy a:xfrm when that is what the part holds', async () => {
    const opened = await slideWithFrame(tableFrameXml('a:xfrm', 914400, 1828800, 5486400, 1828800));
    const table = opened.slides.get(0).tables[0];
    expect(table.x).toBeCloseTo(1);
    expect(table.y).toBeCloseTo(2);
    expect(table.width).toBeCloseTo(6);
    expect(table.height).toBeCloseTo(2);
  });

  it('writes p:xfrm for new tables', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTable([['a', 'b'], ['c', 'd']], { x: 1.5, y: 0.5, width: 6, height: 2 });
    const frame = await frameOf(pptx);
    expect(frame).toContain('<p:xfrm>');
    expect(frame).not.toContain('<a:xfrm');
    expect(/<p:xfrm[^>]*><a:off x="1371600" y="457200"/.test(frame)).toBe(true);
  });

  it('mutating a p:xfrm table keeps exactly one transform, before a:graphic', async () => {
    const pptx = await slideWithFrame(tableFrameXml('p:xfrm', 2 * EMU, 3 * EMU, 5486400, 1828800));
    const table = pptx.slides.get(0).tables[0];
    table.x = 3;
    expect(table.x).toBeCloseTo(3);
    expectSingleTransform(await frameOf(pptx));
    const reopened = await Presentation.open(await pptx.toBuffer());
    const fresh = reopened.slides.get(0).tables[0];
    expect(fresh.x).toBeCloseTo(3);
    expect(fresh.y).toBeCloseTo(3);
    expect(fresh.width).toBeCloseTo(6);
    expect(fresh.height).toBeCloseTo(2);
    expect(fresh.data).toEqual([['a', 'b'], ['c', 'd']]);
  });

  it('addRow on a p:xfrm table keeps exactly one transform with the new extent', async () => {
    const pptx = await slideWithFrame(tableFrameXml('p:xfrm', 2 * EMU, 3 * EMU, 5486400, 1828800));
    const table = pptx.slides.get(0).tables[0];
    table.addRow(['e', 'f']);
    expect(table.rows).toBe(3);
    expectSingleTransform(await frameOf(pptx));
    const reopened = await Presentation.open(await pptx.toBuffer());
    const fresh = reopened.slides.get(0).tables[0];
    expect(fresh.rows).toBe(3);
    expect(fresh.data).toEqual([['a', 'b'], ['c', 'd'], ['e', 'f']]);
    expect(fresh.x).toBeCloseTo(2);
    expect(fresh.y).toBeCloseTo(3);
    expect(fresh.height).toBeCloseTo(3);
  });

  it('editing a legacy a:xfrm table reuses that node instead of adding p:xfrm', async () => {
    const pptx = await slideWithFrame(tableFrameXml('a:xfrm', 914400, 1828800, 5486400, 1828800));
    const table = pptx.slides.get(0).tables[0];
    table.y = 4;
    const frame = await frameOf(pptx);
    expectSingleTransform(frame);
    expect(frame).not.toContain('<p:xfrm');
    const reopened = await Presentation.open(await pptx.toBuffer());
    expect(reopened.slides.get(0).tables[0].y).toBeCloseTo(4);
    expect(reopened.slides.get(0).tables[0].x).toBeCloseTo(1);
  });
});
