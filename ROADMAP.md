# Roadmap toward python-pptx parity

`nodejs-pptx` is currently an MVP, not a full `python-pptx` equivalent. It covers creating/opening PPTX files, adding blank slides, text boxes and PNG/JPEG pictures, enumerating basic shapes, reading and editing text runs, paragraph and text-frame alignment and insets, replacing text, and saving while preserving unrelated ZIP parts.

## Architecture status

The codebase is moving away from a single-file implementation. Current modules are organized around package/OOXML concerns, presentations, slides, shapes, and text:

```txt
src/
  package/
    content-types.ts
    pptx-package.ts
    relationships.ts
  ooxml/
    constants.ts
    ids.ts
    paths.ts
    xml.ts
  presentation/
    presentation.ts
    slide-collection.ts
  slides/
    slide.ts
  shapes/
    shape.ts
    shape-collection.ts
    text-box.ts
    picture.ts
  text/
    text-frame.ts
    paragraph.ts
    run.ts
```

## Current scope

| Area | Current state | Needed to approach `python-pptx` |
| --- | --- | --- |
| Slides | Add and retrieve slides, allocate safe slide part names | Delete, duplicate, reorder, richer layouts, masters, and backgrounds |
| Shapes | Shape collection, text boxes, text-bearing autoshapes, generic read-only shapes | Rectangles, circles, lines, connectors, groups, freeforms, geometry editing |
| Text | `TextFrame -> Paragraph -> Run` hierarchy, `slide.text`, `getTextBoxes()`, `addTextBox()`, and cross-run `replaceText()`, plus paragraph and text-frame alignment (`paragraph.alignment`, `textFrame.alignment`: `left`, `center`, `right`, `justify`, mapped to `a:pPr/@algn`), text-frame insets (`textFrame.margin`, `marginLeft`/`marginTop`/`marginRight`/`marginBottom` in inches, reporting the OOXML defaults 0.1" left/right and 0.05" top/bottom), and run-level `run.text` and `run.font` | Lists and bullets beyond the `• ` prefix accepted by `addTextBox()`, autofit, richer regex capture expansion, and full run formatting beyond `run.font` (see 'Text model notes') |
| Images | PNG/JPEG insertion via `slide.addPicture()` from a Buffer, a Uint8Array (including `byteOffset` views), or a filesystem path, positioned and sized with `{x,y,width,height,name}`, with magic-byte format detection, natural pixel size read from the PNG `IHDR` and JPEG `SOFn` headers, and rejection of truncated PNG/JPEG payloads before the package is touched; readback via `getPictures()`/`shape.hasImage` plus `imageFileName`, `imageContentType`, `imageFormat`, and `imageRelationshipId` (issue #21) | Replace, crop, transform, and other insertion formats (GIF, BMP, TIFF, WEBP, SVG, EMF, WMF); readback already recognizes gif, bmp, and tiff content types by extension |
| Tables | Not implemented | Create and edit rows, columns, cells, merges, and styles |
| Charts | Not implemented | Create and modify charts, series, categories, axes, legends, and embedded Excel data |
| Placeholders | Not implemented | Titles, content, pictures, and inheritance from layouts |
| Notes | Not implemented | Read, create, and modify speaker notes |
| Properties | Title/author only at creation | Read and modify all document properties |
| OOXML API | Internal XML model with typed facades | More typed element classes and package-part abstractions |

## Text model notes

Text now exposes the intended hierarchy:

```txt
TextFrame -> Paragraph -> Run
```

`replaceText()` can match text split across multiple runs within the same paragraph, for example `"Informe " + "2026"`. The replacement is written to the first affected run to preserve that run's formatting, and subsequent affected runs are cleared or trimmed. Regex support is intentionally conservative: matching across runs is supported within a paragraph, but replacement capture expansion is not yet modeled.

Paragraph and text-frame alignment and the text-frame insets are already modeled: `paragraph.alignment` and the aggregated `textFrame.alignment` read and write `a:pPr/@algn` as `left`, `center`, `right`, or `justify`, `textFrame.margin` with `marginLeft`/`marginTop`/`marginRight`/`marginBottom` read and write the `a:bodyPr` `lIns`/`tIns`/`rIns`/`bIns` insets in inches and report the OOXML defaults when the attributes are absent, and `addTextBox()` accepts the same `alignment` and `margin*` options. What is still missing on this axis is the rest of the run and paragraph formatting model: `run.font` covers `name`, `size`, `bold`, `italic`, and `color`, but bullets and lists beyond the `• ` prefix that `addTextBox()` recognizes, autofit, and richer capture expansion are not implemented.

## Recommended order

1. Continue hardening the base: package parts, relationships, content types, identifiers, and round-trip preservation.
2. Expand the shape and text models: `slide.shapes`, generic shapes, paragraphs, runs, and formatting.
3. Implement real slide operations: layouts, masters, placeholders, delete, move, and duplicate.
4. Extend images (issue #21 shipped PNG/JPEG insertion, format detection, and readback; replace, crop, and transform remain) and add tables.
5. Add charts and speaker notes.
6. Build a compatibility corpus with PPTX files from PowerPoint, LibreOffice, and Google Slides, including open-save-open round-trip tests.

## Current restrictions

Picture support covers PNG and JPEG insertion only (issue #21): other formats (GIF, BMP, TIFF, WEBP, SVG, EMF, WMF) are rejected with an error naming the detected format before the package is modified, and truncated PNG/JPEG payloads are rejected the same way, since `detectImage()` requires a full PNG chunk chain and a JPEG `EOI` marker. Replace, crop, and transform operations are still missing. Readback is broader than insertion: `imageContentType` resolves gif, bmp, and tiff media by extension, while `imageFormat` reports only `png` and `jpeg`. On the text side, alignment and text-frame insets ship, but bullets and lists beyond the `• ` prefix accepted by `addTextBox()`, autofit, and run formatting beyond `run.font` do not. Do not add tables, charts, notes, animations, SmartArt, audio, or video until the shape/text abstractions and round-trip behavior are stronger.
