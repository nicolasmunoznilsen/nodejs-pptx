import { Presentation } from '../src/index.js';

const pptx = Presentation.create({ title: 'nodejs-pptx basic example', author: 'nodejs-pptx' });
const slide = pptx.slides.add();
slide.addTextBox('Hello from nodejs-pptx', { x: 0.8, y: 0.8, width: 8.4, height: 0.8, fontSize: 32, color: '1A365D' });
slide.addTextBox('Open, read, modify, and save .pptx files.', { x: 0.8, y: 1.8, width: 8.4, height: 0.6, fontSize: 20, alignment: 'center' });
const body = slide.addTextBox('Left aligned.\nCentered.', { x: 0.8, y: 2.6, width: 8.4, height: 1, name: 'Body' });
const bodyShape = slide.shapes.findById(body.id)!;
bodyShape.textFrame.paragraphs[0].alignment = 'left';
bodyShape.textFrame.paragraphs[1].alignment = 'justify';
await pptx.save('dist/basic.pptx');
