/**
 * PPTX/OOXML Constants
 */

// XML Namespaces
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

// Relationship Types
const REL_TYPES = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
  viewProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
  tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
};

// Content Types
const CONTENT_TYPES = {
  presentation: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  slideMaster: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  presProps: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
  viewProps: 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
  tableStyles: 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
};

// EMU conversions (English Metric Units)
const EMU = {
  INCH: 914400,
  CM: 360000,
  PT: 12700,
  PX: 9525, // at 96 DPI
  // Standard slide size 16:9 (13.333" x 7.5")
  SLIDE_WIDTH_16_9: 12192000,
  SLIDE_HEIGHT_16_9: 6858000,
  // Standard slide size 4:3 (10" x 7.5")
  SLIDE_WIDTH_4_3: 9144000,
  SLIDE_HEIGHT_4_3: 6858000,
};

// Placeholder types
const PLACEHOLDER_TYPES = {
  title: 'title',
  body: 'body',
  ctrTitle: 'ctrTitle',
  subTitle: 'subTitle',
  dt: 'dt',
  ftr: 'ftr',
  sldNum: 'sldNum',
  chart: 'chart',
  tbl: 'tbl',
  clipArt: 'clipArt',
  dgm: 'dgm',
  media: 'media',
  sldImg: 'sldImg',
  pic: 'pic',
};

// Standard layout names
const LAYOUT_NAMES = {
  1: 'Title Slide',
  2: 'Title and Content',
  3: 'Section Header',
  4: 'Two Content',
  5: 'Comparison',
  6: 'Title Only',
  7: 'Blank',
  8: 'Content with Caption',
  9: 'Picture with Caption',
  10: 'Title and Vertical Text',
  11: 'Vertical Title and Text',
};

module.exports = {
  NS,
  REL_TYPES,
  CONTENT_TYPES,
  EMU,
  PLACEHOLDER_TYPES,
  LAYOUT_NAMES,
};
