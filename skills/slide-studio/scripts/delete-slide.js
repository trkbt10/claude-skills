#!/usr/bin/env node
/**
 * Deletes a slide from the presentation
 * Usage: node delete-slide.js <working-directory> --slide <number>
 */

const fs = require('fs');
const path = require('path');
const { parseRelationships, generateRelationships, removeRelationship } = require('./lib/relationships');
const { parseContentTypes, generateContentTypes, removeSlideContentType } = require('./lib/content-types');
const { listSlides } = require('./list-slides');

function parseArgs(args) {
  const result = { slide: null };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--slide' && args[i + 1]) {
      result.slide = parseInt(args[i + 1], 10);
      i++;
    }
  }
  return result;
}

function deleteSlide(workDir, slidePosition) {
  const absWork = path.resolve(workDir);
  const slidesDir = path.join(absWork, 'ppt', 'slides');
  const slideRelsDir = path.join(absWork, 'ppt', 'slides', '_rels');
  const presentationPath = path.join(absWork, 'ppt', 'presentation.xml');
  const presentationRelsPath = path.join(absWork, 'ppt', '_rels', 'presentation.xml.rels');
  const contentTypesPath = path.join(absWork, '[Content_Types].xml');

  // Get current slides
  const slides = listSlides(workDir);
  const slideToDelete = slides.find(s => s.position === slidePosition);

  if (!slideToDelete) {
    throw new Error(`Slide at position ${slidePosition} not found`);
  }

  const { slideNum, rId } = slideToDelete;

  // Delete slide XML file
  const slideFile = path.join(slidesDir, `slide${slideNum}.xml`);
  if (fs.existsSync(slideFile)) {
    fs.unlinkSync(slideFile);
  }

  // Delete slide relationships file
  const slideRelsFile = path.join(slideRelsDir, `slide${slideNum}.xml.rels`);
  if (fs.existsSync(slideRelsFile)) {
    fs.unlinkSync(slideRelsFile);
  }

  // Update presentation.xml.rels
  const presRelsXml = fs.readFileSync(presentationRelsPath, 'utf-8');
  const presRels = parseRelationships(presRelsXml);
  removeRelationship(presRels, rId);
  fs.writeFileSync(presentationRelsPath, generateRelationships(presRels));

  // Update presentation.xml - remove sldId entry
  let presXml = fs.readFileSync(presentationPath, 'utf-8');
  const sldIdPattern = new RegExp(`\\s*<p:sldId[^>]*r:id="${rId}"[^>]*/>`, 'g');
  presXml = presXml.replace(sldIdPattern, '');
  fs.writeFileSync(presentationPath, presXml);

  // Update [Content_Types].xml
  const ctXml = fs.readFileSync(contentTypesPath, 'utf-8');
  const ct = parseContentTypes(ctXml);
  removeSlideContentType(ct, slideNum);
  fs.writeFileSync(contentTypesPath, generateContentTypes(ct));

  console.log(`Deleted slide ${slidePosition} (slide${slideNum}.xml)`);
  return { slideNum, position: slidePosition };
}

// CLI execution
if (require.main === module) {
  const args = process.argv.slice(2);
  const workDir = args[0];

  if (!workDir || workDir.startsWith('--')) {
    console.error('Usage: node delete-slide.js <working-directory> --slide <position>');
    process.exit(1);
  }

  const options = parseArgs(args.slice(1));

  if (options.slide === null) {
    console.error('Error: --slide <position> is required');
    process.exit(1);
  }

  try {
    deleteSlide(workDir, options.slide);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { deleteSlide };
