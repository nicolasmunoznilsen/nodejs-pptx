# Roadmap toward python-pptx parity

`nodejs-pptx` is currently an MVP, not a full `python-pptx` equivalent. It covers creating/opening PPTX files, adding blank slides and text boxes, enumerating basic shapes, reading and editing text runs, replacing text, and saving while preserving unrelated ZIP parts.

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
| Text | `TextFrame -> Paragraph -> Run`, basic run font editing, cross-run `replaceText` | Full paragraph formatting, margins, autofit, hyperlinks, robust rich-text rewrite |
| Images | Not implemented | Add, read, replace, crop, and preserve relationships/media |
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

## Recommended order

1. Continue hardening the base: package parts, relationships, content types, identifiers, and round-trip preservation.
2. Expand the shape and text models: `slide.shapes`, generic shapes, paragraphs, runs, and formatting.
3. Implement real slide operations: layouts, masters, placeholders, delete, move, and duplicate.
4. Add images and tables.
5. Add charts and speaker notes.
6. Build a compatibility corpus with PPTX files from PowerPoint, LibreOffice, and Google Slides, including open-save-open round-trip tests.

## Current restrictions

Do not add images, tables, charts, notes, animations, SmartArt, audio, or video until the shape/text abstractions and round-trip behavior are stronger.
