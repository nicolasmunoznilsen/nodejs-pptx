import { TextFrame } from './text-frame.js';

/** Text search: a literal substring or a regular expression (always matched globally). */
export type TextSearch = string | RegExp;

export interface FindTextOptions {
  /** Match case-insensitively. Applies to string searches; for a regular expression, add the `i` flag. */
  caseSensitive?: boolean;
  /** Return at most the first match of each paragraph instead of every occurrence. */
  firstOnly?: boolean;
}

/**
 * Where a match was found. `slideIndex` and `shapeIndex` are positions inside the slide and shape
 * collections, `paragraphIndex` and `runIndex` inside the owning text frame, and `start`/`end` are
 * offsets into `matchedText`, so several matches inside the same run stay distinguishable.
 */
export interface TextMatchLocation {
  /** 0-based position of the slide in `presentation.slides`. */
  slideIndex: number;
  /** 1-based slide number, for humans and log messages. */
  slideNumber: number;
  /** Slide part path inside the package, e.g. `ppt/slides/slide2.xml`. */
  slidePath: string;
  /** Position of the shape in `slide.shapes.all`. */
  shapeIndex: number;
  /** `p:cNvPr/@id` of the shape. */
  shapeId: string;
  /** `p:cNvPr/@name` of the shape. */
  shapeName: string;
  /** Shape kind, e.g. `textBox` or `autoShape`. */
  shapeType: string;
  /** Index of the paragraph inside the shape's text frame. */
  paragraphIndex: number;
  /** Index of the first run touched by the match inside that paragraph. */
  runIndex: number;
  /** How many runs the match spans (more than one when the text is split across runs). */
  runCount: number;
  /** Start offset inside `matchedText`. */
  start: number;
  /** End offset inside `matchedText`. */
  end: number;
  /** The matched text itself. */
  matchedText: string;
  /** The text of the whole paragraph that contains the match. */
  paragraphText: string;
  /** The text of the first run touched by the match. */
  runText: string;
  /** The whole text of the shape that contains the match. */
  shapeText: string;
}

/**
 * A match plus the live shape and text frame it was found in. `shape` is populated by the shape,
 * slide, and presentation scopes; {@link findTextInFrame} has no shape to report and leaves it
 * `undefined`.
 */
export interface TextMatch extends TextMatchLocation {
  /** The matched text, as a convenience alias of {@link TextMatchLocation.matchedText}. */
  text: string;
  /** The shape the match was found in. */
  shape: ShapeLike;
  /** The text frame the match was found in. */
  frame: TextFrame;
}

/** Structural view of a shape, so the search works with shapes, placeholders, and future types. */
export interface ShapeLike {
  readonly id: string;
  readonly name: string;
  readonly type?: string;
  readonly hasTextFrame: boolean;
  readonly textFrame: TextFrame;
}

/** Structural view of a slide, so the search works with slides and any future slide-like object. */
export interface SlideLike {
  readonly path: string;
  readonly shapes: { readonly all: readonly ShapeLike[] };
}

/**
 * Finds every occurrence of `search` inside a text frame, reporting the paragraph, run, and
 * character offsets of each hit. Text split across runs of the same paragraph is matched as a
 * single string, so `"Informe 2026"` is found even when it is stored as `"Informe "` + `"2026"`.
 */
export function findTextInFrame(frame: TextFrame, search: TextSearch, options: FindTextOptions = {}): TextMatch[] {
  const regex = toSearchRegex(search, options);
  const out: TextMatch[] = [];
  frame.paragraphs.forEach((paragraph, paragraphIndex) => {
    const runs = paragraph.runs;
    const endsAt: number[] = [];
    for (const run of runs) endsAt.push((endsAt.at(-1) ?? 0) + run.text.length);
    const paragraphText = paragraph.text;
    if (!paragraphText) return;
    regex.lastIndex = 0;
    let match = regex.exec(paragraphText);
    while (match) {
      const start = match.index;
      const end = start + match[0].length;
      const runIndex = endsAt.findIndex(endAt => endAt > start);
      out.push({
        slideIndex: 0,
        slideNumber: 0,
        slidePath: '',
        shapeIndex: 0,
        shapeId: '',
        shapeName: '',
        shapeType: '',
        paragraphIndex,
        runIndex,
        runCount: runIndex < 0 ? 0 : endsAt.filter(endAt => endAt > start && endAt < end).length + 1,
        start,
        end,
        matchedText: match[0],
        text: match[0],
        paragraphText,
        runText: runIndex < 0 ? '' : runs[runIndex].text,
        shapeText: '',
        shape: undefined as unknown as ShapeLike,
        frame,
      });
      if (options.firstOnly) break;
      if (regex.lastIndex <= start) regex.lastIndex = start + 1;
      match = regex.exec(paragraphText);
    }
  });
  return out;
}

/** Convenience wrapper that searches the first text frame of a shape. */
export function findTextInShape(shape: ShapeLike, search: TextSearch, options?: FindTextOptions): TextMatch[] {
  if (!shape.hasTextFrame) return [];
  return decorateShape(findTextInFrame(shape.textFrame, search, options), shape);
}

/** Searches a list of shapes in document order, filling in the shape location of each match. */
export function findTextInShapes(shapes: readonly ShapeLike[], search: TextSearch, options: FindTextOptions = {}): TextMatch[] {
  const out: TextMatch[] = [];
  shapes.forEach((shape, shapeIndex) => {
    if (!shape.hasTextFrame) return;
    for (const match of decorateShape(findTextInFrame(shape.textFrame, search, options), shape)) {
      match.shapeIndex = shapeIndex;
      out.push(match);
    }
  });
  return out;
}

/** Searches every text-bearing shape of a slide, in document order. */
export function findTextInSlide(slide: SlideLike, search: TextSearch, options: FindTextOptions = {}, slideIndex = 0): TextMatch[] {
  return findTextInShapes(slide.shapes.all, search, options).map(match => ({
    ...match,
    slideIndex,
    slideNumber: slideIndex + 1,
    slidePath: slide.path,
  }));
}

/** Searches a list of slides in presentation order, filling in the slide location of each match. */
export function findTextInSlides(slides: readonly SlideLike[], search: TextSearch, options: FindTextOptions = {}): TextMatch[] {
  const out: TextMatch[] = [];
  slides.forEach((slide, slideIndex) => out.push(...findTextInSlide(slide, search, options, slideIndex)));
  return out;
}

/** Builds the global regular expression used by the search helpers. */
export function toSearchRegex(search: TextSearch, options: FindTextOptions = {}): RegExp {
  if (typeof search === 'string') {
    if (!search) throw new Error('Cannot search for an empty string.');
    return new RegExp(escapeRegExp(search), options.caseSensitive === false ? 'gi' : 'g');
  }
  // Always clone: a shared regular expression carries `lastIndex`, and the search must not leak
  // that state into (or read it from) the caller's object.
  return new RegExp(search.source, search.global ? search.flags : `${search.flags}g`);
}

function decorateShape(matches: TextMatch[], shape: ShapeLike): TextMatch[] {
  const shapeText = shape.textFrame.text;
  for (const match of matches) {
    match.shapeId = shape.id;
    match.shapeName = shape.name;
    match.shapeType = shape.type ?? 'unknown';
    match.shapeText = shapeText;
    match.shape = shape;
  }
  return matches;
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
