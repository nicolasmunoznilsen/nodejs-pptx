import { Presentation } from '../src/index.js';

// Find every occurrence of a string across all slides, with its location.
// Run: npx vite-node examples/find-text.ts input.pptx "search term"
const [input, needle] = process.argv.slice(2);
if (!input || !needle) {
  console.error('Usage: npx vite-node examples/find-text.ts <input.pptx> <search>');
  process.exit(1);
}

const pptx = await Presentation.open(input);
const matches = pptx.findText(needle, { caseSensitive: false });

console.log(`${input}: ${pptx.slides.length} slide(s), ${matches.length} match(es) for "${needle}"`);

for (const match of matches) {
  // Slide order comes from the presentation relationships, not from file names.
  console.log(`\nslide ${match.slideNumber} (${match.slidePath})`);
  console.log(`  shape "${match.shapeName}" (id ${match.shapeId}, index ${match.shapeIndex}, ${match.shapeType})`);
  console.log(`  paragraph ${match.paragraphIndex}, run ${match.runIndex} of the match (${match.runCount} run(s) spanned)`);
  console.log(`  characters ${match.start}-${match.end}: "${match.matchedText}"`);
  console.log(`  context: ${match.paragraphText}`);
}
