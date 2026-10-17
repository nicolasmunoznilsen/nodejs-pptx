
export { RelationshipCollection, type Relationship } from './package/relationships.js';
export { resolveRelationshipTarget, resolveRootRelationshipTarget, relsPath, relativePackageTarget } from './ooxml/paths.js';
export { ContentTypes } from './package/content-types.js';
export { Presentation, inches, type PresentationOptions } from './presentation/presentation.js';
export { SlideCollection } from './presentation/slide-collection.js';
export { Slide } from './slides/slide.js';
export { Shape, type ShapeType } from './shapes/shape.js';
export { ShapeCollection } from './shapes/shape-collection.js';
export { type TextBoxInfo, type TextBoxOptions } from './shapes/text-box.js';
export { type ImageData, type ImageFormat, type PictureInfo, type PictureOptions, contentTypeForExtension, detectImage, naturalImageSize, pictureXml } from './shapes/picture.js';
export { TextFrame } from './text/text-frame.js';
export { Paragraph } from './text/paragraph.js';
export { type ParagraphAlignment, ALIGNMENT_TO_OOXML, OOXML_TO_ALIGNMENT, alignmentToOoxml, alignmentFromOoxml } from './text/alignment.js';
export { type TextFrameMarginSide, MARGIN_ATTR, MARGIN_DEFAULT_EMU, MARGIN_OPTION, MARGIN_SIDES, marginAttrFor, readMarginInches, writeMarginInches, assertMarginSide, assertMarginInches } from './text/margins.js';
export { Run, type RunFont } from './text/run.js';
export { PptxPackage } from './package/pptx-package.js';
export { SlideLayout, SlideLayoutCollection } from './presentation/slide-layout.js';
export { Placeholder, PlaceholderCollection, type PlaceholderType } from './slides/placeholders.js';
