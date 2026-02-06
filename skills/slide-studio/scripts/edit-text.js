#!/usr/bin/env node
/**
 * Edits text content in a slide
 * Usage: node edit-text.js <working-directory> --slide <number> --placeholder <type> --text "New text"
 *        node edit-text.js <working-directory> --slide <number> --shape-id <id> --text "New text"
 */

const fs = require('fs');
const path = require('path');
const { escapeXml } = require('./lib/xml-utils');
const { listSlides } = require('./list-slides');

function parseArgs(args) {
  const result = {
    slide: null,
    placeholder: null,
    shapeId: null,
    text: null,
  };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--slide' && args[i + 1]) {
      result.slide = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--placeholder' && args[i + 1]) {
      result.placeholder = args[i + 1];
      i++;
    } else if (args[i] === '--shape-id' && args[i + 1]) {
      result.shapeId = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--text' && args[i + 1]) {
      result.text = args[i + 1];
      i++;
    }
  }
  return result;
}

function createTextBody(text) {
  // Handle multiline text by creating multiple paragraphs
  const lines = text.split('\n');
  let paragraphs = '';

  for (const line of lines) {
    paragraphs += `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`;
  }

  return paragraphs;
}

function editTextByPlaceholder(slideXml, phType, text) {
  // Find shape with placeholder type and update its text
  // This is complex because the XML structure is nested

  // Strategy: Find the p:sp element that contains the placeholder type,
  // then replace all a:t content within its p:txBody

  const escapedText = escapeXml(text);

  // Pattern to find shape with specific placeholder and its text body
  const shapePattern = new RegExp(
    `(<p:sp[^>]*>[\\s\\S]*?<p:ph[^>]*type="${phType}"[^>]*/>[\\s\\S]*?<p:txBody>[\\s\\S]*?)(<a:p>[\\s\\S]*?</a:p>)([\\s\\S]*?</p:txBody>[\\s\\S]*?</p:sp>)`,
    'g'
  );

  let modified = false;
  let result = slideXml.replace(shapePattern, (match, before, textPart, after) => {
    modified = true;
    // Replace with simple paragraph structure
    const newText = `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapedText}</a:t></a:r></a:p>`;
    return before + newText + after;
  });

  if (!modified) {
    // Try alternate pattern where ph type might be in different position
    const altPattern = new RegExp(
      `(<p:sp>\\s*<p:nvSpPr>\\s*<p:cNvPr[^>]*/>\\s*<p:cNvSpPr[^>]*(?:/>|>[\\s\\S]*?</p:cNvSpPr>)\\s*<p:nvPr>\\s*<p:ph[^>]*type="${phType}"[^>]*/>[\\s\\S]*?</p:nvPr>\\s*</p:nvSpPr>[\\s\\S]*?<p:txBody>)([\\s\\S]*?)(</p:txBody>)`,
      'g'
    );

    result = slideXml.replace(altPattern, (match, before, content, after) => {
      modified = true;
      const newText = `<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapedText}</a:t></a:r></a:p>`;
      return before + newText + after;
    });
  }

  if (!modified) {
    throw new Error(`Placeholder type '${phType}' not found in slide`);
  }

  return result;
}

function editTextByShapeId(slideXml, shapeId, text) {
  const escapedText = escapeXml(text);

  // Find shape by ID and update its text
  const pattern = new RegExp(
    `(<p:sp>\\s*<p:nvSpPr>\\s*<p:cNvPr[^>]*id="${shapeId}"[^>]*/>[\\s\\S]*?<p:txBody>)([\\s\\S]*?)(</p:txBody>)`,
    'g'
  );

  let modified = false;
  const result = slideXml.replace(pattern, (match, before, content, after) => {
    modified = true;
    const newText = `<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapedText}</a:t></a:r></a:p>`;
    return before + newText + after;
  });

  if (!modified) {
    throw new Error(`Shape with id '${shapeId}' not found in slide`);
  }

  return result;
}

function editText(workDir, slidePosition, options) {
  const absWork = path.resolve(workDir);

  // Get slide info
  const slides = listSlides(workDir);
  const slide = slides.find(s => s.position === slidePosition);

  if (!slide) {
    throw new Error(`Slide at position ${slidePosition} not found`);
  }

  const slidePath = path.join(absWork, 'ppt', 'slides', `slide${slide.slideNum}.xml`);

  if (!fs.existsSync(slidePath)) {
    throw new Error(`Slide file not found: ${slidePath}`);
  }

  let slideXml = fs.readFileSync(slidePath, 'utf-8');

  if (options.placeholder) {
    slideXml = editTextByPlaceholder(slideXml, options.placeholder, options.text);
  } else if (options.shapeId !== null) {
    slideXml = editTextByShapeId(slideXml, options.shapeId, options.text);
  } else {
    throw new Error('Either --placeholder or --shape-id must be specified');
  }

  fs.writeFileSync(slidePath, slideXml);
  console.log(`Updated text in slide ${slidePosition}`);
}

// CLI execution
if (require.main === module) {
  const args = process.argv.slice(2);
  const workDir = args[0];

  if (!workDir || workDir.startsWith('--')) {
    console.error('Usage: node edit-text.js <working-directory> --slide <number> --placeholder <type> --text "..."');
    console.error('       node edit-text.js <working-directory> --slide <number> --shape-id <id> --text "..."');
    console.error('');
    console.error('Placeholder types: title, ctrTitle, subTitle, body, ftr, dt, sldNum');
    process.exit(1);
  }

  const options = parseArgs(args.slice(1));

  if (options.slide === null) {
    console.error('Error: --slide is required');
    process.exit(1);
  }

  if (options.text === null) {
    console.error('Error: --text is required');
    process.exit(1);
  }

  if (options.placeholder === null && options.shapeId === null) {
    console.error('Error: Either --placeholder or --shape-id must be specified');
    process.exit(1);
  }

  try {
    editText(workDir, options.slide, options);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { editText };
