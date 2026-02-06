#!/usr/bin/env node
/**
 * Extracts PPTX to a working directory
 * Usage: node unpack.js <input.pptx> <output-directory>
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function unpack(inputPptx, outputDir) {
  // Validate input file
  const absInput = path.resolve(inputPptx);
  if (!fs.existsSync(absInput)) {
    throw new Error(`Input file not found: ${absInput}`);
  }

  // Create output directory
  const absOutput = path.resolve(outputDir);
  fs.mkdirSync(absOutput, { recursive: true });

  // Read and extract ZIP
  const data = fs.readFileSync(absInput);
  const zip = await JSZip.loadAsync(data);

  // Extract all files
  const files = Object.keys(zip.files);
  for (const filePath of files) {
    const zipEntry = zip.files[filePath];

    if (zipEntry.dir) {
      fs.mkdirSync(path.join(absOutput, filePath), { recursive: true });
    } else {
      const content = await zipEntry.async('nodebuffer');
      const fullPath = path.join(absOutput, filePath);
      fs.mkdirSync(path.dirname(fullPath), { recursive: true });
      fs.writeFileSync(fullPath, content);
    }
  }

  console.log(`Unpacked to: ${absOutput}`);
  return absOutput;
}

// CLI execution
if (require.main === module) {
  const [, , inputPptx, outputDir] = process.argv;

  if (!inputPptx || !outputDir) {
    console.error('Usage: node unpack.js <input.pptx> <output-directory>');
    process.exit(1);
  }

  unpack(inputPptx, outputDir)
    .then(() => process.exit(0))
    .catch((err) => {
      console.error('Error:', err.message);
      process.exit(1);
    });
}

module.exports = { unpack };
