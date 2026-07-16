# Roadmap toward python-pptx parity

`nodejs-pptx` is currently an MVP, not a full `python-pptx` equivalent. It covers roughly 10-15% of the practical API surface: creating/opening PPTX files, adding blank slides and text boxes, reading slide text, replacing text, and saving while preserving unrelated ZIP parts.

## Current scope

| Area | Current state | Needed to approach `python-pptx` |
| --- | --- | --- |
| Slides | Add and retrieve slides | Delete, duplicate, reorder, layouts, masters, and backgrounds |
| Shapes | Text boxes only | Generic shape API, rectangles, circles, lines, connectors, groups, and freeforms |
| Text | Simple text with size, color, and font | Paragraphs, runs, bold, italic, alignment, spacing, bullets, margins, autofit, and hyperlinks |
| Images | Not implemented | Add, read, replace, crop, and preserve relationships/media |
| Tables | Not implemented | Create and edit rows, columns, cells, merges, and styles |
| Charts | Not implemented | Create and modify charts, series, categories, axes, legends, and embedded Excel data |
| Placeholders | Not implemented | Titles, content, pictures, and inheritance from layouts |
| Notes | Not implemented | Read, create, and modify speaker notes |
| Properties | Title/author only at creation | Read and modify all document properties |
| OOXML API | Generic `Record<string, unknown>` model | Typed classes for elements, relationships, and package parts |

## Priority technical issue

The implementation still lives mostly in a single `src/index.ts` file and manipulates complete XML objects through `fast-xml-parser`. That is acceptable for the MVP, but a robust library needs module boundaries around package parts, relationships, content types, slides, shapes, and text before adding larger features.

A target structure for the next phase is:

```txt
src/
  package/
  ooxml/
  presentation/
  slides/
  shapes/
  text/
  images/
  tables/
  charts/
```

## Text model gap

`replaceText()` currently replaces each `a:t` node independently. A visually continuous phrase can be split across multiple runs, for example:

```txt
"Informe " + "2026"
```

Searching for `"Informe 2026"` may therefore fail. The next text layer should model the same hierarchy exposed by `python-pptx`:

```txt
TextFrame -> Paragraph -> Run
```

That model should become the basis for multi-run search/replace and formatting.

## Recommended order

1. Build a robust base: split OOXML/package concerns, relationships, content types, identifiers, and round-trip preservation.
2. Add the shape and text models: `slide.shapes`, generic shapes, paragraphs, runs, and formatting.
3. Implement real slide operations: layouts, masters, placeholders, delete, move, and duplicate.
4. Add images and tables.
5. Add charts and speaker notes.
6. Build a compatibility corpus with PPTX files from PowerPoint, LibreOffice, and Google Slides, including open-save-open round-trip tests.

## Next version focus

The next version should prioritize the shape model and text model before tables or charts. Without those abstractions, every new feature will keep manipulating XML directly, making the project harder to maintain as the API grows.
