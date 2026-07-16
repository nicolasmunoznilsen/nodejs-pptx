import { Presentation } from '../src/index.js';

const pptx = Presentation.create({ title: 'nodejs-pptx basic example', author: 'nodejs-pptx' });
const slide = pptx.slides.add();
slide.addTextBox('Hello from nodejs-pptx', { x: 0.8, y: 0.8, width: 8.4, height: 0.8, fontSize: 32, color: '1A365D' });
slide.addTextBox('Open, read, modify, and save .pptx files.', { x: 0.8, y: 1.8, width: 8.4, height: 0.6, fontSize: 20 });
await pptx.save('dist/basic.pptx');
