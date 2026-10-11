import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/index.js';

// Guards the pattern documented in README.md ("Read text from all slides") and examples/read-all-text.ts.
describe('reading text from all slides', () => {
  it('reads every slide in presentation order with slide.text', () => {
    const pptx = Presentation.create({ title: 'Deck' });
    pptx.slides.add().addTextBox('Slide one title', { name: 'Title 1' });
    pptx.slides.add().addTextBox('Slide two title');
    const emptySlide = pptx.slides.add();

    const collected: string[] = [];
    for (const [index, slide] of pptx.slides.toArray().entries()) {
      for (const line of slide.text.split('\n')) {
        if (line.trim()) collected.push(`slide ${index + 1}: ${line}`);
      }
    }

    expect(collected).toEqual(['slide 1: Slide one title', 'slide 2: Slide two title']);
    expect(emptySlide.text).toBe('');
    expect(pptx.slides.toArray()).toHaveLength(3);
  });

  it('attributes each paragraph to its shape through shapes and textFrame', async () => {
    const created = Presentation.create();
    const slide = created.slides.add();
    slide.addTextBox('Hello world', { name: 'Greeting' });
    slide.addTextBox('Second paragraph', { name: 'Body' });
    const pptx = await Presentation.open(await created.toBuffer());

    const entries = pptx.slides.toArray().flatMap((opened, slideIndex) =>
      opened.shapes.all
        .filter(shape => shape.hasTextFrame)
        .flatMap(shape =>
          shape.textFrame.paragraphs
            .map(paragraph => paragraph.text)
            .filter(text => text.trim() !== '')
            .map(text => ({ slide: slideIndex + 1, shape: shape.name, text })),
        ),
    );

    expect(entries).toEqual([
      { slide: 1, shape: 'Greeting', text: 'Hello world' },
      { slide: 1, shape: 'Body', text: 'Second paragraph' },
    ]);
  });
});
