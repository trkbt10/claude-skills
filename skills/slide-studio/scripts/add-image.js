#!/usr/bin/env node
/**
 * Adds an image to a slide
 * Usage: node add-image.js <working-directory> --slide <number> --image <path> [--x <emu>] [--y <emu>] [--width <emu>] [--height <emu>]
 */

const fs = require('fs');
const path = require('path');
const { parseRelationships, generateRelationships, addImageRelationship } = require('./lib/relationships');
const { parseContentTypes, generateContentTypes, ensureImageType } = require('./lib/content-types');
const { EMU } = require('./lib/constants');
const { getNextId } = require('./lib/xml-utils');
const { listSlides } = require('./list-slides');

function parseArgs(args) {
  const result = {
    slide: null,
    image: null,
    x: EMU.INCH, // 1 inch from left
    y: EMU.INCH, // 1 inch from top
    width: null,
    height: null,
  };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--slide' && args[i + 1]) {
      result.slide = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--image' && args[i + 1]) {
      result.image = args[i + 1];
      i++;
    } else if (args[i] === '--x' && args[i + 1]) {
      result.x = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--y' && args[i + 1]) {
      result.y = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--width' && args[i + 1]) {
      result.width = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--height' && args[i + 1]) {
      result.height = parseInt(args[i + 1], 10);
      i++;
    }
  }
  return result;
}

function getImageDimensions(imagePath) {
  // Read image header to get dimensions (PNG and JPEG support)
  const buffer = fs.readFileSync(imagePath);
  const ext = path.extname(imagePath).toLowerCase();

  if (ext === '.png') {
    // PNG: width at bytes 16-19, height at bytes 20-23 (big-endian)
    if (buffer.slice(0, 8).toString('hex') === '89504e470d0a1a0a') {
      const width = buffer.readUInt32BE(16);
      const height = buffer.readUInt32BE(20);
      return { width, height };
    }
  } else if (ext === '.jpg' || ext === '.jpeg') {
    // JPEG: Find SOF0 marker (0xFFC0) for dimensions
    let i = 2;
    while (i < buffer.length - 1) {
      if (buffer[i] === 0xff) {
        const marker = buffer[i + 1];
        if (marker === 0xc0 || marker === 0xc2) {
          const height = buffer.readUInt16BE(i + 5);
          const width = buffer.readUInt16BE(i + 7);
          return { width, height };
        }
        if (marker === 0xd9) break; // EOI
        if (marker >= 0xd0 && marker <= 0xd9) {
          i += 2;
        } else {
          const len = buffer.readUInt16BE(i + 2);
          i += len + 2;
        }
      } else {
        i++;
      }
    }
  }

  // Default dimensions if unable to read
  return { width: 400, height: 300 };
}

function createPicElement(rId, shapeId, shapeName, x, y, cx, cy) {
  return `
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="${shapeId}" name="${shapeName}"/>
          <p:cNvPicPr>
            <a:picLocks noChangeAspect="1"/>
          </p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="${rId}"/>
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="${x}" y="${y}"/>
            <a:ext cx="${cx}" cy="${cy}"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
      </p:pic>`;
}

function addImage(workDir, slidePosition, options) {
  const absWork = path.resolve(workDir);
  const mediaDir = path.join(absWork, 'ppt', 'media');
  const contentTypesPath = path.join(absWork, '[Content_Types].xml');

  // Validate image file
  const imagePath = path.resolve(options.image);
  if (!fs.existsSync(imagePath)) {
    throw new Error(`Image file not found: ${imagePath}`);
  }

  // Get slide info
  const slides = listSlides(workDir);
  const slide = slides.find(s => s.position === slidePosition);

  if (!slide) {
    throw new Error(`Slide at position ${slidePosition} not found`);
  }

  const slideNum = slide.slideNum;
  const slidePath = path.join(absWork, 'ppt', 'slides', `slide${slideNum}.xml`);
  const slideRelsPath = path.join(absWork, 'ppt', 'slides', '_rels', `slide${slideNum}.xml.rels`);

  // Create media directory if needed
  fs.mkdirSync(mediaDir, { recursive: true });

  // Copy image to media folder with unique name
  const ext = path.extname(imagePath);
  const existingMedia = fs.existsSync(mediaDir) ? fs.readdirSync(mediaDir) : [];
  const imageNum = existingMedia.filter(f => f.startsWith('image')).length + 1;
  const mediaFileName = `image${imageNum}${ext}`;
  const mediaPath = path.join(mediaDir, mediaFileName);
  fs.copyFileSync(imagePath, mediaPath);

  // Get image dimensions
  const dimensions = getImageDimensions(imagePath);
  const width = options.width || dimensions.width * EMU.PX;
  const height = options.height || dimensions.height * EMU.PX;

  // Update slide relationships
  let slideRelsXml = '';
  let slideRels = [];
  if (fs.existsSync(slideRelsPath)) {
    slideRelsXml = fs.readFileSync(slideRelsPath, 'utf-8');
    slideRels = parseRelationships(slideRelsXml);
  } else {
    slideRels = [];
  }

  const imageRId = addImageRelationship(slideRels, `../media/${mediaFileName}`);
  fs.writeFileSync(slideRelsPath, generateRelationships(slideRels));

  // Update slide XML
  let slideXml = fs.readFileSync(slidePath, 'utf-8');
  const nextId = getNextId(slideXml);
  const picElement = createPicElement(imageRId, nextId, `Picture ${nextId}`, options.x, options.y, width, height);

  // Insert picture before closing spTree tag
  slideXml = slideXml.replace(/<\/p:spTree>/, `${picElement}\n    </p:spTree>`);
  fs.writeFileSync(slidePath, slideXml);

  // Update [Content_Types].xml to ensure image type is declared
  const ctXml = fs.readFileSync(contentTypesPath, 'utf-8');
  const ct = parseContentTypes(ctXml);
  ensureImageType(ct, ext);
  fs.writeFileSync(contentTypesPath, generateContentTypes(ct));

  console.log(`Added image to slide ${slidePosition}: ${mediaFileName}`);
  return { mediaFileName, rId: imageRId };
}

// CLI execution
if (require.main === module) {
  const args = process.argv.slice(2);
  const workDir = args[0];

  if (!workDir || workDir.startsWith('--')) {
    console.error('Usage: node add-image.js <working-directory> --slide <number> --image <path> [options]');
    console.error('');
    console.error('Options:');
    console.error('  --x <emu>      X position in EMU (914400 = 1 inch)');
    console.error('  --y <emu>      Y position in EMU');
    console.error('  --width <emu>  Width in EMU');
    console.error('  --height <emu> Height in EMU');
    process.exit(1);
  }

  const options = parseArgs(args.slice(1));

  if (options.slide === null) {
    console.error('Error: --slide is required');
    process.exit(1);
  }

  if (options.image === null) {
    console.error('Error: --image is required');
    process.exit(1);
  }

  try {
    addImage(workDir, options.slide, options);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { addImage };
