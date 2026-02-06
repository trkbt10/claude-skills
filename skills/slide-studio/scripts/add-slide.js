#!/usr/bin/env node
/**
 * Adds a new slide to the presentation
 * Usage: node add-slide.js <working-directory> --layout <index> [--position <pos>]
 */

const fs = require('fs');
const path = require('path');
const { parseRelationships, generateRelationships, addRelationship } = require('./lib/relationships');
const { parseContentTypes, generateContentTypes, addSlideContentType } = require('./lib/content-types');
const { REL_TYPES, CONTENT_TYPES } = require('./lib/constants');
const { xmlDeclaration } = require('./lib/xml-utils');

function parseArgs(args) {
  const result = { layout: 2, position: null };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--layout' && args[i + 1]) {
      result.layout = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--position' && args[i + 1]) {
      result.position = parseInt(args[i + 1], 10);
      i++;
    }
  }
  return result;
}

function getNextSlideNumber(slidesDir) {
  if (!fs.existsSync(slidesDir)) {
    return 1;
  }
  const files = fs.readdirSync(slidesDir).filter(f => f.match(/^slide\d+\.xml$/));
  if (files.length === 0) return 1;
  const nums = files.map(f => parseInt(f.match(/\d+/)[0], 10));
  return Math.max(...nums) + 1;
}

function getNextSlideId(presentationXml) {
  const ids = [];
  const pattern = /<p:sldId[^>]*id="(\d+)"/g;
  let match;
  while ((match = pattern.exec(presentationXml)) !== null) {
    ids.push(parseInt(match[1], 10));
  }
  // PPTX slide IDs typically start at 256
  return ids.length > 0 ? Math.max(...ids) + 1 : 256;
}

/**
 * Extract placeholder shapes from a layout XML
 */
function extractPlaceholders(layoutXml) {
  const placeholders = [];
  // Match <p:sp>...</p:sp> elements that contain <p:ph
  const spPattern = /<p:sp>[\s\S]*?<\/p:sp>/g;
  let match;

  while ((match = spPattern.exec(layoutXml)) !== null) {
    const spXml = match[0];
    // Check if this shape has a placeholder
    if (spXml.includes('<p:ph')) {
      // Extract placeholder type and idx
      const phMatch = spXml.match(/<p:ph([^>]*)\/?>/);
      if (phMatch) {
        const phAttrs = phMatch[1];
        const typeMatch = phAttrs.match(/type="([^"]*)"/);
        const idxMatch = phAttrs.match(/idx="([^"]*)"/);
        const szMatch = phAttrs.match(/sz="([^"]*)"/);

        placeholders.push({
          type: typeMatch ? typeMatch[1] : null,
          idx: idxMatch ? idxMatch[1] : null,
          sz: szMatch ? szMatch[1] : null,
        });
      }
    }
  }

  return placeholders;
}

/**
 * Generate placeholder shape XML for a slide
 */
function generatePlaceholderShape(id, name, phType, phIdx, phSz, sampleText) {
  const phAttrs = [];
  if (phType) phAttrs.push(`type="${phType}"`);
  if (phSz) phAttrs.push(`sz="${phSz}"`);
  if (phIdx) phAttrs.push(`idx="${phIdx}"`);

  const phAttrStr = phAttrs.length > 0 ? ' ' + phAttrs.join(' ') : '';

  return `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${id}" name="${name}"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph${phAttrStr}/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${sampleText}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>`;
}

/**
 * Get sample text for a placeholder type
 */
function getSampleText(phType, layoutName) {
  switch (phType) {
    case 'title':
    case 'ctrTitle':
      return layoutName || 'Title';
    case 'subTitle':
      return 'Subtitle';
    case 'body':
      return 'Body text';
    case 'dt':
      return '';  // Date - usually auto-filled
    case 'ftr':
      return '';  // Footer
    case 'sldNum':
      return '';  // Slide number - auto-filled
    case 'pic':
      return '';  // Picture placeholder
    default:
      return 'Content';
  }
}

function createSlideXml(layoutNum, placeholders, layoutName) {
  let shapesXml = '';
  let shapeId = 2;

  // Skip footer-type placeholders (dt, ftr, sldNum) for cleaner output
  const contentPlaceholders = placeholders.filter(
    ph => !['dt', 'ftr', 'sldNum'].includes(ph.type)
  );

  for (const ph of contentPlaceholders) {
    const name = ph.type ? `${ph.type} ${shapeId - 1}` : `Shape ${shapeId - 1}`;
    const sampleText = getSampleText(ph.type, layoutName);
    if (sampleText) {  // Only add if there's sample text
      shapesXml += generatePlaceholderShape(shapeId, name, ph.type, ph.idx, ph.sz, sampleText) + '\n';
      shapeId++;
    }
  }

  return `${xmlDeclaration()}<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
${shapesXml}    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
}

function createSlideRelsXml(layoutNum) {
  return `${xmlDeclaration()}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideLayout}" Target="../slideLayouts/slideLayout${layoutNum}.xml"/>
</Relationships>`;
}

function addSlide(workDir, layoutIndex, position = null) {
  const absWork = path.resolve(workDir);
  const slidesDir = path.join(absWork, 'ppt', 'slides');
  const slideRelsDir = path.join(absWork, 'ppt', 'slides', '_rels');
  const presentationPath = path.join(absWork, 'ppt', 'presentation.xml');
  const presentationRelsPath = path.join(absWork, 'ppt', '_rels', 'presentation.xml.rels');
  const contentTypesPath = path.join(absWork, '[Content_Types].xml');

  // Validate layout exists
  const layoutPath = path.join(absWork, 'ppt', 'slideLayouts', `slideLayout${layoutIndex}.xml`);
  if (!fs.existsSync(layoutPath)) {
    throw new Error(`Layout ${layoutIndex} not found: ${layoutPath}`);
  }

  // Read layout XML and extract placeholders
  const layoutXml = fs.readFileSync(layoutPath, 'utf-8');
  const placeholders = extractPlaceholders(layoutXml);

  // Extract layout name from cSld name attribute
  const nameMatch = layoutXml.match(/<p:cSld[^>]*name="([^"]*)"/);
  const layoutName = nameMatch ? nameMatch[1] : `Layout ${layoutIndex}`;

  // Create directories if needed
  fs.mkdirSync(slidesDir, { recursive: true });
  fs.mkdirSync(slideRelsDir, { recursive: true });

  // Get next slide number
  const slideNum = getNextSlideNumber(slidesDir);

  // Create slide XML with placeholders from layout
  const slideXml = createSlideXml(layoutIndex, placeholders, layoutName);
  fs.writeFileSync(path.join(slidesDir, `slide${slideNum}.xml`), slideXml);

  // Create slide relationships
  const slideRelsXml = createSlideRelsXml(layoutIndex);
  fs.writeFileSync(path.join(slideRelsDir, `slide${slideNum}.xml.rels`), slideRelsXml);

  // Update presentation.xml.rels
  const presRelsXml = fs.readFileSync(presentationRelsPath, 'utf-8');
  const presRels = parseRelationships(presRelsXml);
  const slideRId = addRelationship(presRels, REL_TYPES.slide, `slides/slide${slideNum}.xml`);
  fs.writeFileSync(presentationRelsPath, generateRelationships(presRels));

  // Update presentation.xml
  let presXml = fs.readFileSync(presentationPath, 'utf-8');
  const slideId = getNextSlideId(presXml);

  // Add sldId to sldIdLst
  const sldIdEntry = `<p:sldId id="${slideId}" r:id="${slideRId}"/>`;

  if (position !== null) {
    // Insert at specific position
    const sldIdPattern = /<p:sldId[^>]*\/>/g;
    const matches = [...presXml.matchAll(sldIdPattern)];

    if (position <= 1 || matches.length === 0) {
      // Insert at beginning
      presXml = presXml.replace(/<p:sldIdLst>/, `<p:sldIdLst>\n    ${sldIdEntry}`);
    } else if (position > matches.length) {
      // Insert at end
      presXml = presXml.replace(/<\/p:sldIdLst>/, `  ${sldIdEntry}\n  </p:sldIdLst>`);
    } else {
      // Insert at position
      const targetMatch = matches[position - 1];
      presXml = presXml.slice(0, targetMatch.index) + sldIdEntry + '\n    ' + presXml.slice(targetMatch.index);
    }
  } else {
    // Append at end
    presXml = presXml.replace(/<\/p:sldIdLst>/, `  ${sldIdEntry}\n  </p:sldIdLst>`);
  }

  fs.writeFileSync(presentationPath, presXml);

  // Update [Content_Types].xml
  const ctXml = fs.readFileSync(contentTypesPath, 'utf-8');
  const ct = parseContentTypes(ctXml);
  addSlideContentType(ct, slideNum);
  fs.writeFileSync(contentTypesPath, generateContentTypes(ct));

  console.log(`Added slide ${slideNum} using layout ${layoutIndex}`);
  return { slideNum, slideId, rId: slideRId };
}

// CLI execution
if (require.main === module) {
  const args = process.argv.slice(2);
  const workDir = args[0];

  if (!workDir || workDir.startsWith('--')) {
    console.error('Usage: node add-slide.js <working-directory> --layout <index> [--position <pos>]');
    process.exit(1);
  }

  const options = parseArgs(args.slice(1));

  try {
    addSlide(workDir, options.layout, options.position);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { addSlide };
