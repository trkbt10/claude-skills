#!/usr/bin/env node
/**
 * Visualizes slide structure as ASCII art
 * Helps LLM understand shape positions and sizes
 *
 * Usage: node visualize-slide.js <working-directory> [--slide <n>]
 */

const fs = require('fs');
const path = require('path');

// Standard slide size (16:9 widescreen)
const SLIDE_WIDTH = 12192000;  // EMU
const SLIDE_HEIGHT = 6858000;  // EMU

// ASCII grid size
const GRID_WIDTH = 72;
const GRID_HEIGHT = 24;

/**
 * Convert EMU to grid position
 */
function emuToGrid(emuX, emuY) {
  const x = Math.round((emuX / SLIDE_WIDTH) * GRID_WIDTH);
  const y = Math.round((emuY / SLIDE_HEIGHT) * GRID_HEIGHT);
  return { x: Math.min(x, GRID_WIDTH - 1), y: Math.min(y, GRID_HEIGHT - 1) };
}

/**
 * Convert EMU to inches for display
 */
function emuToInches(emu) {
  return (emu / 914400).toFixed(2);
}

/**
 * Extract shapes from slide XML
 */
function extractShapes(slideXml) {
  const shapes = [];

  // Match <p:sp> elements
  const spPattern = /<p:sp>[\s\S]*?<\/p:sp>/g;
  let match;
  let id = 1;

  while ((match = spPattern.exec(slideXml)) !== null) {
    const spXml = match[0];

    // Extract name
    const nameMatch = spXml.match(/<p:cNvPr[^>]*name="([^"]*)"/);
    const name = nameMatch ? nameMatch[1] : `Shape ${id}`;

    // Extract position
    const offMatch = spXml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const x = offMatch ? parseInt(offMatch[1], 10) : 0;
    const y = offMatch ? parseInt(offMatch[2], 10) : 0;

    // Extract size
    const extMatch = spXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    const width = extMatch ? parseInt(extMatch[1], 10) : 0;
    const height = extMatch ? parseInt(extMatch[2], 10) : 0;

    // Extract placeholder type if present
    const phMatch = spXml.match(/<p:ph[^>]*type="([^"]*)"/);
    const phType = phMatch ? phMatch[1] : null;

    // Extract text content
    const textMatches = spXml.match(/<a:t>([^<]*)<\/a:t>/g);
    const text = textMatches
      ? textMatches.map(t => t.replace(/<\/?a:t>/g, '')).join(' ').substring(0, 30)
      : '';

    shapes.push({
      id: id++,
      name,
      x, y, width, height,
      phType,
      text: text.trim(),
    });
  }

  return shapes;
}

/**
 * Render ASCII grid
 */
function renderGrid(shapes) {
  // Initialize grid
  const grid = Array(GRID_HEIGHT).fill(null).map(() => Array(GRID_WIDTH).fill(' '));

  // Draw border
  for (let x = 0; x < GRID_WIDTH; x++) {
    grid[0][x] = '-';
    grid[GRID_HEIGHT - 1][x] = '-';
  }
  for (let y = 0; y < GRID_HEIGHT; y++) {
    grid[y][0] = '|';
    grid[y][GRID_WIDTH - 1] = '|';
  }
  grid[0][0] = '+';
  grid[0][GRID_WIDTH - 1] = '+';
  grid[GRID_HEIGHT - 1][0] = '+';
  grid[GRID_HEIGHT - 1][GRID_WIDTH - 1] = '+';

  // Draw shapes
  for (const shape of shapes) {
    const start = emuToGrid(shape.x, shape.y);
    const end = emuToGrid(shape.x + shape.width, shape.y + shape.height);

    // Clamp to grid bounds (inside border)
    const x1 = Math.max(1, Math.min(start.x, GRID_WIDTH - 2));
    const y1 = Math.max(1, Math.min(start.y, GRID_HEIGHT - 2));
    const x2 = Math.max(1, Math.min(end.x, GRID_WIDTH - 2));
    const y2 = Math.max(1, Math.min(end.y, GRID_HEIGHT - 2));

    // Draw shape boundary
    for (let x = x1; x <= x2; x++) {
      if (y1 > 0 && y1 < GRID_HEIGHT - 1) grid[y1][x] = '-';
      if (y2 > 0 && y2 < GRID_HEIGHT - 1) grid[y2][x] = '-';
    }
    for (let y = y1; y <= y2; y++) {
      if (x1 > 0 && x1 < GRID_WIDTH - 1) grid[y][x1] = '|';
      if (x2 > 0 && x2 < GRID_WIDTH - 1) grid[y][x2] = '|';
    }

    // Draw corners
    if (y1 > 0 && x1 > 0) grid[y1][x1] = '+';
    if (y1 > 0 && x2 < GRID_WIDTH - 1) grid[y1][x2] = '+';
    if (y2 < GRID_HEIGHT - 1 && x1 > 0) grid[y2][x1] = '+';
    if (y2 < GRID_HEIGHT - 1 && x2 < GRID_WIDTH - 1) grid[y2][x2] = '+';

    // Draw shape ID in center
    const centerX = Math.floor((x1 + x2) / 2);
    const centerY = Math.floor((y1 + y2) / 2);
    const label = `[${shape.id}]`;
    if (centerY > 0 && centerY < GRID_HEIGHT - 1) {
      for (let i = 0; i < label.length && centerX + i < GRID_WIDTH - 1; i++) {
        grid[centerY][centerX + i] = label[i];
      }
    }
  }

  return grid.map(row => row.join('')).join('\n');
}

/**
 * Generate shape legend
 */
function generateLegend(shapes) {
  const lines = ['', 'Shapes:'];
  lines.push('─'.repeat(70));

  for (const shape of shapes) {
    const pos = `(${emuToInches(shape.x)}", ${emuToInches(shape.y)}")`;
    const size = `${emuToInches(shape.width)}" × ${emuToInches(shape.height)}"`;
    const type = shape.phType ? `[${shape.phType}]` : '';
    const text = shape.text ? `"${shape.text}..."` : '';

    lines.push(`[${shape.id}] ${shape.name} ${type}`);
    lines.push(`    Position: ${pos}  Size: ${size}`);
    if (text) lines.push(`    Text: ${text}`);
  }

  lines.push('─'.repeat(70));
  lines.push(`Slide: ${emuToInches(SLIDE_WIDTH)}" × ${emuToInches(SLIDE_HEIGHT)}" (16:9)`);

  return lines.join('\n');
}

function visualizeSlide(workDir, slideNum = 1) {
  const absWork = path.resolve(workDir);
  const slidePath = path.join(absWork, 'ppt', 'slides', `slide${slideNum}.xml`);

  if (!fs.existsSync(slidePath)) {
    throw new Error(`Slide not found: ${slidePath}`);
  }

  const slideXml = fs.readFileSync(slidePath, 'utf-8');
  const shapes = extractShapes(slideXml);

  console.log(`\nSlide ${slideNum} Layout:`);
  console.log(renderGrid(shapes));
  console.log(generateLegend(shapes));

  return shapes;
}

// CLI execution
if (require.main === module) {
  const args = process.argv.slice(2);
  let workDir = null;
  let slideNum = 1;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--slide' && args[i + 1]) {
      slideNum = parseInt(args[i + 1], 10);
      i++; // Skip next arg
    } else if (!args[i].startsWith('--')) {
      workDir = args[i];
    }
  }

  if (!workDir) {
    console.error('Usage: node visualize-slide.js <working-directory> [--slide <n>]');
    process.exit(1);
  }

  try {
    visualizeSlide(workDir, slideNum);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { visualizeSlide, extractShapes };
