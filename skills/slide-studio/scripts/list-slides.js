#!/usr/bin/env node
/**
 * Lists all slides in a presentation with their titles
 * Usage: node list-slides.js <working-directory>
 */

const fs = require('fs');
const path = require('path');
const { parseRelationships } = require('./lib/relationships');

function listSlides(workDir) {
  const absWork = path.resolve(workDir);
  const presentationPath = path.join(absWork, 'ppt', 'presentation.xml');
  const presentationRelsPath = path.join(absWork, 'ppt', '_rels', 'presentation.xml.rels');

  if (!fs.existsSync(presentationPath)) {
    throw new Error(`presentation.xml not found: ${presentationPath}`);
  }

  const presentationXml = fs.readFileSync(presentationPath, 'utf-8');
  const relsXml = fs.existsSync(presentationRelsPath)
    ? fs.readFileSync(presentationRelsPath, 'utf-8')
    : '';

  const relationships = parseRelationships(relsXml);
  const slides = [];

  // Parse sldIdLst to get slide order
  const sldIdPattern = /<p:sldId[^>]*id="(\d+)"[^>]*r:id="([^"]*)"/g;
  let match;
  let position = 1;

  while ((match = sldIdPattern.exec(presentationXml)) !== null) {
    const slideId = match[1];
    const rId = match[2];

    // Find slide file from relationship
    const rel = relationships.find(r => r.id === rId);
    if (!rel) continue;

    const slideFile = rel.target.replace(/^.*\//, ''); // Get filename only
    const slideNum = parseInt(slideFile.match(/\d+/)?.[0] || '0', 10);
    const slidePath = path.join(absWork, 'ppt', 'slides', slideFile);

    let title = '';
    let layoutIndex = null;

    if (fs.existsSync(slidePath)) {
      const slideXml = fs.readFileSync(slidePath, 'utf-8');

      // Try to extract title text
      // Look for title placeholder
      const titleMatch = slideXml.match(/<p:ph[^>]*type="(?:title|ctrTitle)"[^>]*\/>[\s\S]*?<a:t>([^<]*)<\/a:t>/);
      if (titleMatch) {
        title = titleMatch[1];
      } else {
        // Fallback: get first text
        const firstText = slideXml.match(/<a:t>([^<]+)<\/a:t>/);
        if (firstText) {
          title = firstText[1].substring(0, 50);
        }
      }

      // Get layout from slide relationships
      const slideRelsPath = path.join(absWork, 'ppt', 'slides', '_rels', `${slideFile}.rels`);
      if (fs.existsSync(slideRelsPath)) {
        const slideRelsXml = fs.readFileSync(slideRelsPath, 'utf-8');
        const layoutMatch = slideRelsXml.match(/Target="[^"]*slideLayout(\d+)\.xml"/);
        if (layoutMatch) {
          layoutIndex = parseInt(layoutMatch[1], 10);
        }
      }
    }

    slides.push({
      position,
      slideNum,
      slideId,
      rId,
      title: title || '(no title)',
      layoutIndex,
    });

    position++;
  }

  return slides;
}

function printSlides(slides) {
  console.log('Slides:');
  console.log('=======');
  if (slides.length === 0) {
    console.log('  (no slides)');
    return;
  }
  for (const slide of slides) {
    const layout = slide.layoutIndex ? `[layout ${slide.layoutIndex}]` : '';
    console.log(`  ${slide.position}. ${slide.title} ${layout}`);
  }
  console.log(`\nTotal: ${slides.length} slides`);
}

// CLI execution
if (require.main === module) {
  const [, , workDir] = process.argv;

  if (!workDir) {
    console.error('Usage: node list-slides.js <working-directory>');
    process.exit(1);
  }

  try {
    const slides = listSlides(workDir);
    printSlides(slides);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { listSlides };
