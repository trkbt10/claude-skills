/**
 * Relationships management for .rels files
 */

const { REL_TYPES } = require('./constants');
const { xmlDeclaration } = require('./xml-utils');

/**
 * Parse .rels XML into structured data
 */
function parseRelationships(xml) {
  const relationships = [];
  const pattern = /<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"(?:\s+TargetMode="([^"]+)")?\s*\/>/g;
  let match;

  while ((match = pattern.exec(xml)) !== null) {
    relationships.push({
      id: match[1],
      type: match[2],
      target: match[3],
      targetMode: match[4] || null,
    });
  }

  return relationships;
}

/**
 * Generate .rels XML from relationships array
 */
function generateRelationships(relationships) {
  let xml = xmlDeclaration();
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';

  for (const rel of relationships) {
    xml += `  <Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"`;
    if (rel.targetMode) {
      xml += ` TargetMode="${rel.targetMode}"`;
    }
    xml += '/>\n';
  }

  xml += '</Relationships>';
  return xml;
}

/**
 * Get next available rId
 */
function getNextRId(relationships) {
  const ids = relationships
    .map(r => parseInt(r.id.replace('rId', ''), 10))
    .filter(n => !isNaN(n));
  return 'rId' + (ids.length > 0 ? Math.max(...ids) + 1 : 1);
}

/**
 * Add a relationship
 */
function addRelationship(relationships, type, target, targetMode = null) {
  const id = getNextRId(relationships);
  relationships.push({
    id,
    type,
    target,
    targetMode,
  });
  return id;
}

/**
 * Remove a relationship by ID
 */
function removeRelationship(relationships, id) {
  const index = relationships.findIndex(r => r.id === id);
  if (index !== -1) {
    relationships.splice(index, 1);
  }
  return relationships;
}

/**
 * Find relationship by target
 */
function findByTarget(relationships, target) {
  return relationships.find(r => r.target === target || r.target.endsWith(target));
}

/**
 * Find relationship by type
 */
function findByType(relationships, type) {
  return relationships.filter(r => r.type === type);
}

/**
 * Find relationship by ID
 */
function findById(relationships, id) {
  return relationships.find(r => r.id === id);
}

/**
 * Add slide relationship to presentation
 */
function addSlideRelationship(relationships, slideNum) {
  return addRelationship(relationships, REL_TYPES.slide, `slides/slide${slideNum}.xml`);
}

/**
 * Add layout relationship to slide
 */
function addLayoutRelationship(relationships, layoutNum) {
  return addRelationship(relationships, REL_TYPES.slideLayout, `../slideLayouts/slideLayout${layoutNum}.xml`);
}

/**
 * Add image relationship
 */
function addImageRelationship(relationships, imagePath) {
  return addRelationship(relationships, REL_TYPES.image, imagePath);
}

/**
 * Get standard root relationships
 */
function getStandardRootRels() {
  return [
    {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'ppt/presentation.xml',
    },
    {
      id: 'rId2',
      type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
      target: 'docProps/core.xml',
    },
    {
      id: 'rId3',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
      target: 'docProps/app.xml',
    },
  ];
}

/**
 * Update relationship target (for renumbering slides etc.)
 */
function updateTarget(relationships, id, newTarget) {
  const rel = findById(relationships, id);
  if (rel) {
    rel.target = newTarget;
  }
  return relationships;
}

module.exports = {
  parseRelationships,
  generateRelationships,
  getNextRId,
  addRelationship,
  removeRelationship,
  findByTarget,
  findByType,
  findById,
  addSlideRelationship,
  addLayoutRelationship,
  addImageRelationship,
  getStandardRootRels,
  updateTarget,
};
