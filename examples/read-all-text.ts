import { Presentation } from '../src/index.js';

// Read the text of every slide of an existing .pptx file.
// Run: npx vite-node examples/read-all-text.ts input.pptx
const input = process.argv[2];
if (!input) {
  console.error('Usage: npx vite-node examples/read-all-text.ts <input.pptx>');
  process.exit(1);
}

const pptx = await Presentation.open(input);
console.log(`${input}: ${pptx.slides.length} slide(s)`);

// Slide order comes from the presentation relationships, not from file names.
for (const [index, slide] of pptx.slides.toArray().entries()) {
  console.log(`\nSlide ${index + 1} (${slide.path})`);

  // `slide.text` returns every paragraph of every text-bearing shape, one per line, in document order.
  for (const line of slide.text.split('\n')) {
    if (line.trim()) console.log(`  ${line}`);
  }

  // The same text attributed to the shape and paragraph it came from.
  for (const shape of slide.shapes.all) {
    if (!shape.hasTextFrame) continue;
    shape.textFrame.paragraphs.forEach((paragraph, paragraphIndex) => {
      if (!paragraph.text.trim()) return;
      const runs = paragraph.runs.map(run => run.text).join(' | ');
      console.log(`  shape "${shape.name}" (id ${shape.id}) paragraph ${paragraphIndex}: ${paragraph.text} [runs: ${runs}]`);
    });
  }
}
