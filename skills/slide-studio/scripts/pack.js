#!/usr/bin/env node
/**
 * Creates PPTX from a working directory
 * Usage: node pack.js <working-directory> <output.pptx>
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

/**
 * Recursively get all files in a directory
 */
function getAllFiles(dirPath, basePath = dirPath) {
  const files = [];
  const entries = fs.readdirSync(dirPath, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);
    const relativePath = path.relative(basePath, fullPath);

    // Skip macOS system files but keep .rels files (required by OOXML)
    if (entry.name === '.DS_Store') continue;

    if (entry.isDirectory()) {
      files.push(...getAllFiles(fullPath, basePath));
    } else {
      files.push(relativePath);
    }
  }

  return files;
}

async function pack(workDir, outputPptx) {
  const absWork = path.resolve(workDir);
  const absOutput = path.resolve(outputPptx);

  // Validate working directory
  if (!fs.existsSync(absWork)) {
    throw new Error(`Working directory not found: ${absWork}`);
  }

  // Verify essential PPTX files exist
  const contentTypesPath = path.join(absWork, '[Content_Types].xml');
  if (!fs.existsSync(contentTypesPath)) {
    throw new Error('Invalid PPTX structure: [Content_Types].xml not found');
  }

  // Create ZIP
  const zip = new JSZip();

  // Get all files
  const files = getAllFiles(absWork);

  // PPTX requires specific ordering: [Content_Types].xml should be first
  // Sort files to ensure proper order
  const sortedFiles = files.sort((a, b) => {
    // [Content_Types].xml first
    if (a === '[Content_Types].xml') return -1;
    if (b === '[Content_Types].xml') return 1;
    // _rels directory early
    if (a.startsWith('_rels/') && !b.startsWith('_rels/')) return -1;
    if (!a.startsWith('_rels/') && b.startsWith('_rels/')) return 1;
    // Alphabetical otherwise
    return a.localeCompare(b);
  });

  // Add files to ZIP
  for (const file of sortedFiles) {
    const fullPath = path.join(absWork, file);
    const content = fs.readFileSync(fullPath);
    zip.file(file, content);
  }

  // Generate ZIP content
  const zipContent = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  // Remove existing output if present
  if (fs.existsSync(absOutput)) {
    fs.unlinkSync(absOutput);
  }

  // Write output file
  fs.mkdirSync(path.dirname(absOutput), { recursive: true });
  fs.writeFileSync(absOutput, zipContent);

  console.log(`Created: ${absOutput}`);
  return absOutput;
}

// CLI execution
if (require.main === module) {
  const [, , workDir, outputPptx] = process.argv;

  if (!workDir || !outputPptx) {
    console.error('Usage: node pack.js <working-directory> <output.pptx>');
    process.exit(1);
  }

  pack(workDir, outputPptx)
    .then(() => process.exit(0))
    .catch((err) => {
      console.error('Error:', err.message);
      process.exit(1);
    });
}

module.exports = { pack };
