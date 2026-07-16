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

## API implemented in this MVP

- `Presentation.create(options?)`
- `Presentation.open(path | Buffer)`
- `presentation.save(path)`
- `presentation.toBuffer()`
- `presentation.slides.length`
- `presentation.slides.get(index)`
- `presentation.slides.add()`
- `slide.text`
- `slide.getTextBoxes()`
- `slide.replaceText(search, replacement)`
- `slide.addTextBox(text, options)`

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

See `examples/basic.ts` for a complete example.
