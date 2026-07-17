import type { Presentation } from './presentation.js';
import type { Slide } from '../slides/slide.js';

export class SlideCollection {
  constructor(private readonly presentation: Presentation) {}
  get length(): number { return this.presentation.slideList.length; }
  get(index: number): Slide {
    if (!Number.isInteger(index) || index < 0 || index >= this.length) throw new RangeError(`Slide index ${index} is out of range (0-${this.length - 1}).`);
    return this.presentation.slideList[index];
  }
  add(): Slide { return this.presentation.addSlideInternal(); }
  [Symbol.iterator](): IterableIterator<Slide> { return this.presentation.slideList[Symbol.iterator](); }
}
