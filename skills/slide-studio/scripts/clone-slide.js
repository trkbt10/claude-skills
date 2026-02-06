#!/usr/bin/env node
/**
 * Duplicates an existing slide
 * Usage: node clone-slide.js <working-directory> --source <position> [--position <pos>]
 */

const fs = require('fs');
const path = require('path');
const { parseRelationships, generateRelationships, addRelationship } = require('./lib/relationships');
const { parseContentTypes, generateContentTypes, addSlideContentType } = require('./lib/content-types');
const { REL_TYPES } = require('./lib/constants');
const { listSlides } = require('./list-slides');

function parseArgs(args) {
  const result = { source: null, position: null };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--source' && args[i + 1]) {
      result.source = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--position' && args[i + 1]) {
      result.position = parseInt(args[i + 1], 10);
      i++;
    }
  }
  return result;
}

function getNextSlideNumber(slidesDir) {
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
  return ids.length > 0 ? Math.max(...ids) + 1 : 256;
}

function updateSlideIds(slideXml, offset) {
  // Update all id attributes to be unique
  return slideXml.replace(/(\bid\s*=\s*["'])(\d+)(["'])/g, (match, before, id, after) => {
    return before + (parseInt(id, 10) + offset) + after;
  });
}

function cloneSlide(workDir, sourcePosition, targetPosition = null) {
  const absWork = path.resolve(workDir);
  const slidesDir = path.join(absWork, 'ppt', 'slides');
  const slideRelsDir = path.join(absWork, 'ppt', 'slides', '_rels');
  const presentationPath = path.join(absWork, 'ppt', 'presentation.xml');
  const presentationRelsPath = path.join(absWork, 'ppt', '_rels', 'presentation.xml.rels');
  const contentTypesPath = path.join(absWork, '[Content_Types].xml');

  // Get source slide info
  const slides = listSlides(workDir);
  const sourceSlide = slides.find(s => s.position === sourcePosition);

  if (!sourceSlide) {
    throw new Error(`Source slide at position ${sourcePosition} not found`);
  }

  const sourceNum = sourceSlide.slideNum;
  const newNum = getNextSlideNumber(slidesDir);

  // Copy slide XML
  const sourceSlideXml = fs.readFileSync(path.join(slidesDir, `slide${sourceNum}.xml`), 'utf-8');
  const newSlideXml = updateSlideIds(sourceSlideXml, 1000 + newNum * 100); // Offset IDs to avoid conflicts
  fs.writeFileSync(path.join(slidesDir, `slide${newNum}.xml`), newSlideXml);

  // Copy slide relationships
  const sourceRelsPath = path.join(slideRelsDir, `slide${sourceNum}.xml.rels`);
  if (fs.existsSync(sourceRelsPath)) {
    const sourceRelsXml = fs.readFileSync(sourceRelsPath, 'utf-8');
    fs.writeFileSync(path.join(slideRelsDir, `slide${newNum}.xml.rels`), sourceRelsXml);
  }

  // Update presentation.xml.rels
  const presRelsXml = fs.readFileSync(presentationRelsPath, 'utf-8');
  const presRels = parseRelationships(presRelsXml);
  const slideRId = addRelationship(presRels, REL_TYPES.slide, `slides/slide${newNum}.xml`);
  fs.writeFileSync(presentationRelsPath, generateRelationships(presRels));

  // Update presentation.xml
  let presXml = fs.readFileSync(presentationPath, 'utf-8');
  const slideId = getNextSlideId(presXml);
  const sldIdEntry = `<p:sldId id="${slideId}" r:id="${slideRId}"/>`;

  if (targetPosition !== null) {
    const sldIdPattern = /<p:sldId[^>]*\/>/g;
    const matches = [...presXml.matchAll(sldIdPattern)];

    if (targetPosition <= 1 || matches.length === 0) {
      presXml = presXml.replace(/<p:sldIdLst>/, `<p:sldIdLst>\n    ${sldIdEntry}`);
    } else if (targetPosition > matches.length) {
      presXml = presXml.replace(/<\/p:sldIdLst>/, `  ${sldIdEntry}\n  </p:sldIdLst>`);
    } else {
      const targetMatch = matches[targetPosition - 1];
      presXml = presXml.slice(0, targetMatch.index) + sldIdEntry + '\n    ' + presXml.slice(targetMatch.index);
    }
  } else {
    presXml = presXml.replace(/<\/p:sldIdLst>/, `  ${sldIdEntry}\n  </p:sldIdLst>`);
  }

  fs.writeFileSync(presentationPath, presXml);

  // Update [Content_Types].xml
  const ctXml = fs.readFileSync(contentTypesPath, 'utf-8');
  const ct = parseContentTypes(ctXml);
  addSlideContentType(ct, newNum);
  fs.writeFileSync(contentTypesPath, generateContentTypes(ct));

  console.log(`Cloned slide ${sourcePosition} to new slide ${newNum}`);
  return { slideNum: newNum, slideId, rId: slideRId };
}

// CLI execution
if (require.main === module) {
  const args = process.argv.slice(2);
  const workDir = args[0];

  if (!workDir || workDir.startsWith('--')) {
    console.error('Usage: node clone-slide.js <working-directory> --source <position> [--position <pos>]');
    process.exit(1);
  }

  const options = parseArgs(args.slice(1));

  if (options.source === null) {
    console.error('Error: --source <position> is required');
    process.exit(1);
  }

  try {
    cloneSlide(workDir, options.source, options.position);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { cloneSlide };
