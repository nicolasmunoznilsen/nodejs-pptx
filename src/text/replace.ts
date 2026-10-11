/**
 * Cross-run text replacement, shared by every text-bearing container in a deck: slide shapes,
 * slide/layout placeholders, and any future part that exposes a `Run` list.
 *
 * The match is computed over the concatenation of the run texts of a single paragraph, so a
 * phrase split across runs ("Informe " + "2026") is found as one match. The replacement is
 * written into the first affected run to preserve that run's formatting, and later affected runs
 * are cleared or trimmed.
 */

/** Minimal structural view of a `Run`, so callers can pass either `Run` instances or raw adapters. */
export interface ReplacableRun {
  text: string;
}

export interface TextMatch {
  start: number;
  end: number;
}

/** Replace every occurrence of `search` in one paragraph and return how many were replaced. */
export function replaceInRuns(runs: ReplacableRun[], search: string | RegExp, replacement: string): number {
  const full = runs.map(run => run.text).join('');
  const matches = typeof search === 'string' ? findStringMatches(full, search) : findRegexMatches(full, search);
  // Apply from the end so earlier offsets stay valid while run texts change.
  for (const match of matches.reverse()) applyReplacement(runs, match.start, match.end, replacement);
  return matches.length;
}

export function findStringMatches(text: string, needle: string): TextMatch[] {
  if (!needle) return [];
  const matches: TextMatch[] = [];
  let start = text.indexOf(needle);
  while (start >= 0) {
    matches.push({ start, end: start + needle.length });
    start = text.indexOf(needle, start + needle.length);
  }
  return matches;
}

export function findRegexMatches(text: string, regex: RegExp): TextMatch[] {
  const flags = regex.flags.includes('g') ? regex.flags : `${regex.flags}g`;
  const clone = new RegExp(regex.source, flags);
  const matches: TextMatch[] = [];
  for (const match of text.matchAll(clone)) {
    if (match.index === undefined || match[0] === '') continue;
    matches.push({ start: match.index, end: match.index + match[0].length });
  }
  return matches;
}

function applyReplacement(runs: ReplacableRun[], start: number, end: number, replacement: string): void {
  let pos = 0;
  let wrote = false;
  for (const run of runs) {
    const runStart = pos;
    const runEnd = pos + run.text.length;
    pos = runEnd;
    if (runEnd <= start || runStart >= end) continue;
    const before = run.text.slice(0, Math.max(0, start - runStart));
    const after = run.text.slice(Math.max(0, end - runStart));
    if (!wrote) { run.text = before + replacement + after; wrote = true; } else { run.text = after; }
  }
}
