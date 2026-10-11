import { Presentation } from '../src/index.js';

// Replace text across every slide of an existing .pptx file.
// Run: npx vite-node examples/replace-in-presentation.ts input.pptx FY2026 FY2027
const [input, search, replacement] = process.argv.slice(2);
if (!input || search === undefined || replacement === undefined) {
  console.error('Usage: npx vite-node examples/replace-in-presentation.ts <input.pptx> <search> <replacement>');
  process.exit(1);
}

const pptx = await Presentation.open(input);

// Slides are visited in presentation order (p:sldId relationships), including placeholders,
// shapes nested inside groups, and matches that cross run boundaries inside a paragraph.
const replaced = pptx.replaceText(search, replacement);
console.log(`${input}: ${replaced} replacement(s) across ${pptx.slides.length} slide(s)`);

// Layout text is template content, so it is only rewritten when explicitly requested.
// const replaced = pptx.replaceText(search, replacement, { includeLayouts: true });

if (replaced) await pptx.save(input.replace(/\.pptx$/i, '') + '.replaced.pptx');
