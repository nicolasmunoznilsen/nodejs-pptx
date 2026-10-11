import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Presentation, type PlaceholderType } from '../src/index.js';
import { mapPlaceholderType } from '../src/slides/placeholders.js';

/** Builds a deck whose first slide contains one `<p:sp>` with `<p:ph type="...">`. */
async function slideWithPlaceholderType(type: string): Promise<Presentation> {
  const pptx = Presentation.create();
  pptx.slides.add().addTextBox('body', { name: 'Seed' });
  const zip = await JSZip.loadAsync(await pptx.toBuffer());
  const path = 'ppt/slides/slide1.xml';
  const xml = await zip.file(path)!.async('string');
  const shape = `<p:sp><p:nvSpPr><p:cNvPr id="99" name="Injected"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>injected</a:t></a:r></a:p></p:txBody></p:sp>`;
  zip.file(path, xml.replace('</p:spTree>', `${shape}</p:spTree>`));
  return Presentation.open(await zip.generateAsync({ type: 'nodebuffer' }));
}

describe('placeholder type mapping', () => {
  it('maps every OOXML p:ph type spelling', () => {
    expect(mapPlaceholderType('title')).toBe('title');
    expect(mapPlaceholderType('ctrTitle')).toBe('centerTitle');
    expect(mapPlaceholderType('subTitle')).toBe('subtitle');
    expect(mapPlaceholderType('body')).toBe('body');
    expect(mapPlaceholderType('object')).toBe('object');
    expect(mapPlaceholderType('pic')).toBe('picture');
    expect(mapPlaceholderType('dt')).toBe('date');
    expect(mapPlaceholderType('ftr')).toBe('footer');
    expect(mapPlaceholderType('sldNum')).toBe('slideNumber');
  });

  it('accepts the public PlaceholderType spellings written by other tools', () => {
    for (const name of ['title', 'centerTitle', 'subtitle', 'body', 'object', 'picture', 'date', 'footer', 'slideNumber'] as PlaceholderType[]) {
      expect(mapPlaceholderType(name)).toBe(name);
    }
  });

  it('rejects names inherited from Object.prototype instead of resolving them', () => {
    for (const inherited of ['constructor', 'toString', 'valueOf', 'hasOwnProperty', '__proto__', 'prototype']) {
      expect(mapPlaceholderType(inherited)).toBe('unknown');
    }
    // The guard must not be a truthiness check: `constructor` really is resolvable without it.
    expect((({} as Record<string, unknown>)['constructor'])).toBeTypeOf('function');
  });

  it('maps unknown, empty, and malformed names to unknown', () => {
    for (const raw of ['nope', 'Title', 'TITLE', ' ctrTitle', 'ctrTitle ', '', 'null', 'undefined', '0', '1']) {
      expect(mapPlaceholderType(raw)).toBe('unknown');
    }
    expect(mapPlaceholderType(undefined as unknown as string)).toBe('unknown');
    expect(mapPlaceholderType(null as unknown as string)).toBe('unknown');
    expect(mapPlaceholderType({ toString: () => 'title' } as unknown as string)).toBe('unknown');
  });

  it('never returns a non-string for an inherited document type', async () => {
    const opened = await slideWithPlaceholderType('constructor');
    const placeholder = opened.slides.get(0).placeholders.findByName('Injected')!;
    expect(placeholder.type).toBe('unknown');
    expect(typeof placeholder.type).toBe('string');
  });

  it('reads real placeholder types and unknown ones from the document', async () => {
    const known = await slideWithPlaceholderType('ctrTitle');
    expect(known.slides.get(0).placeholders.findByName('Injected')!.type).toBe('centerTitle');
    expect(known.slides.get(0).placeholders.findByType('centerTitle')?.name).toBe('Injected');

    const unknown = await slideWithPlaceholderType('weirdType');
    expect(unknown.slides.get(0).placeholders.findByName('Injected')!.type).toBe('unknown');

    const empty = await slideWithPlaceholderType('');
    expect(empty.slides.get(0).placeholders.findByName('Injected')!.type).toBe('unknown');
  });

  it('keeps reading untouched decks byte-identically', async () => {
    const pptx = Presentation.create();
    pptx.slides.add().addTextBox('read only', { name: 'ReadOnly' });
    const before = await pptx.toBuffer();
    const opened = await Presentation.open(before);
    expect(opened.slides.get(0).placeholders.length).toBe(0);
    expect(await opened.toBuffer()).toEqual(before);
  });
});
