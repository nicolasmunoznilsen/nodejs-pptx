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

## Find text across all slides

`presentation.findText(search)` searches every text-bearing shape of every slide, in presentation order, and returns one entry per match with its full location: slide, shape, paragraph, run, and character offsets. `search` is a string or a regular expression.

```ts
import { Presentation } from 'nodejs-pptx';

const pptx = await Presentation.open('input.pptx');

for (const match of pptx.findText('FY2026')) {
  console.log(
    `slide ${match.slideNumber} (${match.slidePath})`,
    `shape "${match.shapeName}" #${match.shapeId} at ${match.shapeIndex}`,
    `paragraph ${match.paragraphIndex}, run ${match.runIndex}`,
    `chars ${match.start}-${match.end}`,
  );
}
```

Each match carries:

| Field | Meaning |
| --- | --- |
| `slideIndex` / `slideNumber` / `slidePath` | Position in `presentation.slides`, 1-based number, and slide part inside the package |
| `shapeIndex` / `shapeId` / `shapeName` / `shapeType` | Position in `slide.shapes.all` plus `p:cNvPr` identity and kind |
| `paragraphIndex` / `runIndex` / `runCount` | Paragraph and first affected run inside the text frame, and how many runs the match spans |
| `start` / `end` / `matchedText` / `text` | Offsets inside `matchedText` and the matched text itself |
| `paragraphText` / `runText` / `shapeText` | Surrounding text, for previews and context |
| `shape` / `frame` | The live shape and `TextFrame`, so a match can be edited in place |

Text split across runs of the same paragraph is matched as a single string, so `"Informe 2026"` is found even when it is stored as `"Informe "` + `"2026"`; `runIndex` then points at the first affected run and `runCount` is greater than `1`. Matches are reported in document order, and repeated hits inside one run keep their own offsets.

Options:

```ts
pptx.findText('informe', { caseSensitive: false }); // also matches "Informe"
pptx.findText(/FY\d{4}/);                            // regular expressions, matched globally
pptx.findText('hola', { firstOnly: true });          // at most the first match per paragraph
```

Scopes narrower than the whole deck, plus the standalone helpers used by all of them:

```ts
slide.find('FY2026');                 // one slide, same match shape
findTextInShape(shape, 'FY2026');     // one shape
findTextInFrame(shape.textFrame, 'x'); // one text frame
```

`slide.find()` resolves the slide position through its presentation, so `slideIndex` stays correct after `slides.move()` or `slides.duplicate()`. Shapes without a text frame (pictures, and graphic frames until tables and charts are implemented) never match. An empty search string throws.

`examples/find-text.ts` is a runnable version of the snippet:

```bash
npx vite-node examples/find-text.ts input.pptx "search term"
```

## API implemented in this MVP

- `Presentation.create(options?)`
- `Presentation.open(path | Buffer)`
- `presentation.save(path)`
- `presentation.toBuffer()`
- `presentation.findText(search, options?)`
- `presentation.slides.length`
- `presentation.slides.get(index)`
- `presentation.slides.add()`
- `presentation.slides.toArray()`
- `slide.text`
- `slide.find(search, options?)`
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
