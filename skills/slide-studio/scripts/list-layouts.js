#!/usr/bin/env node
/**
 * Lists available slide layouts in a presentation
 * Usage: node list-layouts.js <working-directory>
 */

const fs = require('fs');
const path = require('path');
const { getAllText } = require('./lib/xml-utils');

function listLayouts(workDir) {
  const absWork = path.resolve(workDir);
  const layoutsDir = path.join(absWork, 'ppt', 'slideLayouts');

  if (!fs.existsSync(layoutsDir)) {
    throw new Error(`Layouts directory not found: ${layoutsDir}`);
  }

  const layouts = [];
  const files = fs.readdirSync(layoutsDir)
    .filter(f => f.match(/^slideLayout\d+\.xml$/))
    .sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)[0], 10);
      const numB = parseInt(b.match(/\d+/)[0], 10);
      return numA - numB;
    });

  for (const file of files) {
    const num = parseInt(file.match(/\d+/)[0], 10);
    const xml = fs.readFileSync(path.join(layoutsDir, file), 'utf-8');

    // Try to extract layout name from cSld name attribute
    const nameMatch = xml.match(/<p:cSld[^>]*name="([^"]*)"/);
    let name = nameMatch ? nameMatch[1] : `Layout ${num}`;

    // Get placeholder types
    const phTypes = [];
    const phPattern = /<p:ph[^>]*type="([^"]*)"/g;
    let match;
    while ((match = phPattern.exec(xml)) !== null) {
      if (!phTypes.includes(match[1])) {
        phTypes.push(match[1]);
      }
    }

    layouts.push({
      index: num,
      name,
      file,
      placeholders: phTypes,
    });
  }

  return layouts;
}

function printLayouts(layouts) {
  console.log('Available Layouts:');
  console.log('==================');
  for (const layout of layouts) {
    const placeholders = layout.placeholders.length > 0
      ? `(${layout.placeholders.join(', ')})`
      : '(no placeholders)';
    console.log(`  ${layout.index}: ${layout.name} ${placeholders}`);
  }
}

// CLI execution
if (require.main === module) {
  const [, , workDir] = process.argv;

  if (!workDir) {
    console.error('Usage: node list-layouts.js <working-directory>');
    process.exit(1);
  }

  try {
    const layouts = listLayouts(workDir);
    printLayouts(layouts);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { listLayouts };
