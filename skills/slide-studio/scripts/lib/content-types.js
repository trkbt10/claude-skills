/**
 * Content Types management for [Content_Types].xml
 */

const { CONTENT_TYPES } = require('./constants');
const { xmlDeclaration } = require('./xml-utils');

/**
 * Parse [Content_Types].xml into structured data
 */
function parseContentTypes(xml) {
  const result = {
    defaults: {},
    overrides: {},
  };

  // Parse Default elements
  const defaultPattern = /<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"\s*\/>/g;
  let match;
  while ((match = defaultPattern.exec(xml)) !== null) {
    result.defaults[match[1]] = match[2];
  }

  // Parse Override elements
  const overridePattern = /<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"\s*\/>/g;
  while ((match = overridePattern.exec(xml)) !== null) {
    result.overrides[match[1]] = match[2];
  }

  return result;
}

/**
 * Generate [Content_Types].xml from structured data
 */
function generateContentTypes(data) {
  let xml = xmlDeclaration();
  xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';

  // Add Default elements
  for (const [ext, type] of Object.entries(data.defaults)) {
    xml += `  <Default Extension="${ext}" ContentType="${type}"/>\n`;
  }

  // Add Override elements (sorted for consistency)
  const sortedOverrides = Object.entries(data.overrides).sort((a, b) => a[0].localeCompare(b[0]));
  for (const [partName, type] of sortedOverrides) {
    xml += `  <Override PartName="${partName}" ContentType="${type}"/>\n`;
  }

  xml += '</Types>';
  return xml;
}

/**
 * Add a Default content type for file extension
 */
function addDefault(data, extension, contentType) {
  data.defaults[extension] = contentType;
  return data;
}

/**
 * Add an Override content type for specific part
 */
function addOverride(data, partName, contentType) {
  // Ensure partName starts with /
  const normalizedPath = partName.startsWith('/') ? partName : '/' + partName;
  data.overrides[normalizedPath] = contentType;
  return data;
}

/**
 * Remove an Override by part name
 */
function removeOverride(data, partName) {
  const normalizedPath = partName.startsWith('/') ? partName : '/' + partName;
  delete data.overrides[normalizedPath];
  return data;
}

/**
 * Add slide to content types
 */
function addSlideContentType(data, slideNum) {
  return addOverride(data, `/ppt/slides/slide${slideNum}.xml`, CONTENT_TYPES.slide);
}

/**
 * Remove slide from content types
 */
function removeSlideContentType(data, slideNum) {
  return removeOverride(data, `/ppt/slides/slide${slideNum}.xml`);
}

/**
 * Ensure image extension has content type
 */
function ensureImageType(data, extension) {
  const ext = extension.toLowerCase().replace('.', '');
  if (!data.defaults[ext]) {
    const typeMap = {
      png: CONTENT_TYPES.png,
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: CONTENT_TYPES.gif,
      bmp: 'image/bmp',
      svg: 'image/svg+xml',
    };
    if (typeMap[ext]) {
      data.defaults[ext] = typeMap[ext];
    }
  }
  return data;
}

/**
 * Get standard content types for a new presentation
 */
function getStandardContentTypes() {
  return {
    defaults: {
      rels: 'application/vnd.openxmlformats-package.relationships+xml',
      xml: 'application/xml',
      jpeg: 'image/jpeg',
      png: 'image/png',
    },
    overrides: {
      '/ppt/presentation.xml': CONTENT_TYPES.presentation,
      '/ppt/presProps.xml': CONTENT_TYPES.presProps,
      '/ppt/viewProps.xml': CONTENT_TYPES.viewProps,
      '/ppt/tableStyles.xml': CONTENT_TYPES.tableStyles,
      '/docProps/core.xml': 'application/vnd.openxmlformats-package.core-properties+xml',
      '/docProps/app.xml': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    },
  };
}

module.exports = {
  parseContentTypes,
  generateContentTypes,
  addDefault,
  addOverride,
  removeOverride,
  addSlideContentType,
  removeSlideContentType,
  ensureImageType,
  getStandardContentTypes,
};
