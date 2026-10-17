# nodejs-pptx

`nodejs-pptx` is a TypeScript, ESM-first PowerPoint library for Node.js 20+ inspired by `python-pptx`.

The MVP focuses on the core workflow that generic generators do not cover: open an existing `.pptx`, read the real slide order from OOXML relationships, edit slide text in place, preserve unmodified package parts, and save the file again.

## Install

```bash
npm install nodejs-pptx
```

## Create a presentation

```ts
import { Presentation } from 'nodejs-pptx';

const pptx = Presentation.create({ title: 'Demo', author: 'Acme' });
const slide = pptx.slides.add();
slide.addTextBox('Hello PowerPoint', { x: 1, y: 1, width: 6, height: 1 });
await pptx.save('demo.pptx');
```

## Open, read, edit, and save an existing PPTX

```ts
import { Presentation } from 'nodejs-pptx';

const pptx = await Presentation.open('input.pptx');
const firstSlide = pptx.slides.get(0);

console.log(firstSlide.text);
console.log(firstSlide.getTextBoxes());

firstSlide.replaceText('FY2026', 'FY2027');
firstSlide.addTextBox('Reviewed by Finance', { x: 0.7, y: 4.8, width: 4, height: 0.4 });

await pptx.save('output.pptx');
```

## Read text from all slides

`slide.text` returns the text of every text-bearing shape on a slide, one paragraph per line, in document order. Iterate `pptx.slides` to read the whole deck; the order comes from the `p:sldId` relationships in `ppt/presentation.xml`, not from ZIP file names.

```ts
import { Presentation } from 'nodejs-pptx';

const pptx = await Presentation.open('input.pptx');

for (const [index, slide] of pptx.slides.toArray().entries()) {
  console.log(`Slide ${index + 1}`);
  for (const line of slide.text.split('\n')) {
    if (line.trim()) console.log(`  ${line}`);
  }
}
```

To know which shape and paragraph each line came from, walk `slide.shapes` and the `TextFrame -> Paragraph -> Run` hierarchy instead:

```ts
for (const [index, slide] of pptx.slides.toArray().entries()) {
  for (const shape of slide.shapes.all) {
    if (!shape.hasTextFrame) continue;
    shape.textFrame.paragraphs.forEach((paragraph, paragraphIndex) => {
      if (paragraph.text) console.log(index + 1, shape.name, shape.id, paragraphIndex, paragraph.text);
    });
  }
}
```

Slides without any text-bearing shape yield an empty `slide.text`, so filter empty lines when building a search index or a transcript. Shapes such as pictures and graphic frames (tables and charts) currently contribute no text; see [`ROADMAP.md`](ROADMAP.md).

`examples/read-all-text.ts` is a runnable version of both snippets:

```bash
npx vite-node examples/read-all-text.ts input.pptx
```

## API implemented in this MVP

- `Presentation.create(options?)`
- `Presentation.open(path | Buffer)`
- `presentation.save(path)`
- `presentation.toBuffer()`
- `presentation.slides.length`
- `presentation.slides.get(index)`
- `presentation.slides.add()`
- `presentation.slides.toArray()`
- `slide.text`
- `slide.getTextBoxes()`
- `slide.replaceText(search, replacement)`
- `slide.addTextBox(text, options)`
- `slide.shapes.length`
- `slide.shapes.get(index)`
- `slide.shapes.findByName(name)`
- `slide.shapes.findById(id)`
- `slide.shapes.all`
- `shape.textFrame.paragraphs`
- `paragraph.runs`
- `paragraph.alignment`
- `textFrame.alignment`
- `textFrame.marginLeft` / `marginTop` / `marginRight` / `marginBottom` / `margin`
- `run.text`
- `run.font`

## Shapes API

Slides now expose a `shapes` collection in addition to the legacy text-box helpers:

```ts
const shape = slide.shapes.findByName('Title 1');

if (shape?.hasTextFrame) {
  console.log(shape.id, shape.name, shape.type);
  console.log(shape.x, shape.y, shape.width, shape.height);
}
```

Implemented shape operations are intentionally read-focused for this phase:

- `slide.shapes.length`
- `slide.shapes.get(index)`
- `slide.shapes.findByName(name)`
- `slide.shapes.findById(id)`
- `slide.shapes.all`
- base `Shape` properties: `id`, `name`, `type`, `x`, `y`, `width`, `height`, `hasTextFrame`, `textFrame` (`x`, `y`, `width`, `height` are in inches; the package stores EMU)

Text boxes and text-bearing autoshapes expose `textFrame`; unknown shape types are preserved and surfaced as generic shapes.

## Images API

PNG and JPEG images can be inserted from a `Buffer`, a `Uint8Array` (including a view with a `byteOffset`), or a filesystem path:

```ts
const picture = slide.addPicture('assets/logo.png', { x: 1, y: 2, width: 3, height: 1.5, name: 'Logo' });
slide.addPicture(await readFile('assets/photo.jpg'));
slide.addPicture(new Uint8Array(await readFile('assets/icon.png')), { x: 0.5, y: 0.5 }); // raw bytes from any source
```

`addPicture` detects the format from the magic bytes, writes the exact bytes of the payload to a free `ppt/media/imageN` part (unique across formats, never overwriting an existing part), adds the slide image relationship, registers the `png`/`jpg`/`jpeg` content-type defaults once, and emits a `p:pic` shape. `x`, `y`, `width`, and `height` are inches converted to EMU; omitted `width`/`height` fall back to the natural image size at 96 DPI (640x480 when the header carries no dimensions). It returns `{ name, id, index, x, y, width, height, fileName, contentType, format, relationshipId }`.

Anything else throws `Unsupported image format: <format>` before the package is touched, so a failed call never adds a media part, relationship, shape, or content type. The error names `GIF`, `BMP`, `TIFF`, `WEBP`, `SVG`, `EMF`, `WMF`, `truncated PNG`, `truncated JPEG`, or `unknown`.

Read side:

- `slide.getPictures()` returns the same info objects for existing pictures.
- `shape.type === 'picture'`, `shape.hasImage`, `shape.imageRelationshipId`, `shape.imageFileName`, `shape.imageContentType`, `shape.imageFormat`.
- `detectImage(buffer)` exposes `{ format, extension, contentType, defaultExtensions, buffer, width, height }`.

## TextFrame, Paragraph, and Run API

Text is exposed as a hierarchy similar to `python-pptx`:

```ts
const shape = slide.shapes.findByName('Title 1');
const paragraph = shape?.textFrame.paragraphs[0];
const run = paragraph?.runs[0];

if (run) {
  run.text = 'Nuevo título';
  run.font.bold = true;
  run.font.italic = true;
  run.font.color = '#1F497D';
  run.font.name = 'Calibri';
  run.font.size = 24;
}
```

### Paragraph alignment

`paragraph.alignment` reads and writes the OOXML `a:pPr/@algn` attribute using the public names
`'left'`, `'center'`, `'right'`, and `'justify'` (mapped to `l`, `ctr`, `r`, and `just`):

```ts
const paragraph = slide.shapes.findByName('Title 1')!.textFrame.paragraphs[0];

paragraph.alignment = 'center';   // <a:pPr algn="ctr"/>
paragraph.alignment = undefined;  // removes algn, so the value is inherited again
```

The getter returns `undefined` when the paragraph has no explicit `algn` attribute, or when that
attribute carries a value this library does not understand, which means the alignment is inherited
from the paragraph/list style, the placeholder, or the theme. Assigning any other value throws a
`TypeError`. `textFrame.alignment` is a convenience over the whole frame:
reading returns the shared alignment only when every paragraph agrees (otherwise `undefined`), and
writing applies it to all paragraphs. `slide.addTextBox()` also accepts `alignment`, and
`slide.getTextBoxes()` reports it. Existing `a:pPr` attributes and children (`marL`, `indent`,
`a:buChar`, `a:lnSpc`, ...) are kept; a newly created `a:pPr` is inserted as the first child of
`a:p` as required by the schema. Reading or assigning an unchanged alignment does not mark the
slide dirty, so untouched slides stay byte-identical after a round trip.

`examples/basic.ts` is a runnable version of this snippet.

Both directions of the mapping (`alignmentToOoxml` / `alignmentFromOoxml`, also exported) resolve
through `Object.hasOwn` on frozen lookup tables, so names inherited from `Object.prototype`
(`constructor`, `toString`, `__proto__`, ...) never resolve: a document containing `algn="constructor"`
reads back as inherited alignment (`undefined`) instead of leaking an inherited value, and assigning
such a name throws instead of writing it into the XML.

### Text-frame margins

`textFrame.marginLeft`, `marginTop`, `marginRight`, and `marginBottom` read and write the OOXML
text-frame insets `a:bodyPr/@lIns`, `@tIns`, `@rIns`, and `@bIns`, in inches:

```ts
const frame = slide.shapes.findByName('Title 1')!.textFrame;

frame.marginLeft = 0.25;   // <a:bodyPr lIns="228600"/>
frame.marginTop = 0.1;     // <a:bodyPr ... tIns="91440"/>
frame.marginLeft = undefined; // removes lIns, so the side inherits the default again
```

When an inset is not explicit, PowerPoint applies its DrawingML default: **0.1" left, 0.05" top,
0.1" right, 0.05" bottom** (`MARGIN_DEFAULT_EMU` = `91440 / 45720 / 91440 / 45720` EMU). The
getters report that rendered default rather than `undefined`, and assigning `undefined` means
"inherit default": it deletes the attribute, so the next read returns the default again. Values are
stored as rounded EMU (`Math.round(inches * 914400)`); `NaN`, negative, non-finite, or non-number
values throw a `TypeError`, as does an unknown side name.

`textFrame.margin` is a convenience over the four insets: reading returns the shared value only when
all four agree (otherwise `undefined`), writing applies it to all of them. `slide.addTextBox()` also
accepts `margin` and `marginLeft`/`marginTop`/`marginRight`/`marginBottom`, and
`slide.getTextBoxes()` reports them. Unrelated `a:bodyPr` attributes and children (`wrap`,
`anchor`, `rtlCol`, `a:normAutofit`, `a:sp3d`, ...) are kept, and a missing `a:bodyPr` is created
lazily — only on write — as the first child of `p:txBody`, as required by the schema. Reading a
margin, or assigning a value that is already in effect, does not mark the slide dirty, so untouched
slides stay byte-identical after a round trip.

The side -> attribute and side -> default tables (`MARGIN_ATTR`, `MARGIN_DEFAULT_EMU`, also
exported, plus `marginAttrFor`) resolve through `Object.hasOwn` on frozen objects, so names
inherited from `Object.prototype` (`constructor`, `__proto__`, ...) never resolve: a document
containing `lIns="constructor"` reads back as the default inset instead of leaking an inherited
value, and using such a side throws instead of writing it into the XML.

`slide.replaceText()` can replace text that is split across multiple runs in the same paragraph:

```ts
slide.replaceText('Informe 2026', 'Informe 2027');
```

When a match crosses run boundaries, the replacement is written into the first affected run and later affected runs are cleared or trimmed so the first run's formatting is preserved. Regex replacements are supported for text within a paragraph, but replacement capture expansion is not implemented yet; pass the final replacement string explicitly.

## Internal architecture

The codebase is now split into modules for presentation orchestration, slides, shapes, text, package relationships/content types, and low-level OOXML utilities:

```txt
src/
  package/
  ooxml/
  presentation/
  slides/
  shapes/
  text/
```

Images are limited to PNG and JPEG insertion (no replace, crop, or transform, and no GIF/BMP/TIFF/SVG/EMF/WMF). This phase deliberately does not add tables, charts, notes, animations, SmartArt, audio, or video. The focus is internal robustness, conservative round-tripping, path resolution, IDs, relationships, and a maintainable shape/text model.

## Preservation model

When opening a `.pptx`, the library reads `_rels/.rels`, resolves `ppt/presentation.xml`, reads `ppt/_rels/presentation.xml.rels`, locates the real slide parts referenced by `p:sldId`, and preserves every other ZIP entry unchanged. `replaceText` mutates only `a:t` text nodes in the target slide XML; it does not rebuild unrelated OOXML.

## Not implemented yet

Tables, charts, notes, animations, and SmartArt are intentionally out of scope for this MVP. Image editing beyond PNG/JPEG insertion (replace, crop, other formats) is also still missing.

## Development

```bash
npm run build
npm run lint
npm test
```

See `examples/basic.ts` for a complete create/save example and `examples/read-all-text.ts` for reading text from all slides.

For planned architecture and python-pptx parity work, see [`ROADMAP.md`](ROADMAP.md).
