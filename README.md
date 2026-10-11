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
- `shape.textFrame.paragraphs`
- `paragraph.runs`
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
- base `Shape` properties: `id`, `name`, `type`, `x`, `y`, `width`, `height`, `hasTextFrame`, `textFrame`

Text boxes and text-bearing autoshapes expose `textFrame`; unknown shape types are preserved and surfaced as generic shapes.

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

This phase deliberately does not add images, tables, charts, notes, animations, SmartArt, audio, or video. The focus is internal robustness, conservative round-tripping, path resolution, IDs, relationships, and a maintainable shape/text model.

## Preservation model

When opening a `.pptx`, the library reads `_rels/.rels`, resolves `ppt/presentation.xml`, reads `ppt/_rels/presentation.xml.rels`, locates the real slide parts referenced by `p:sldId`, and preserves every other ZIP entry unchanged. `replaceText` mutates only `a:t` text nodes in the target slide XML; it does not rebuild unrelated OOXML.

## Not implemented yet

Tables, charts, images, notes, animations, and SmartArt are intentionally out of scope for this MVP.

## Development

```bash
npm run build
npm run lint
npm test
```

See `examples/basic.ts` for a complete create/save example and `examples/read-all-text.ts` for reading text from all slides.

For planned architecture and python-pptx parity work, see [`ROADMAP.md`](ROADMAP.md).
