#!/usr/bin/env node
/**
 * Adds a shape (text box, rectangle, etc.) to a slide at custom position
 * Supports theme colors for consistency with applied templates
 *
 * Usage:
 *   node add-shape.js <work-dir> --slide <n> --type textbox --text "Hello"
 *       --x <inches> --y <inches> --width <inches> --height <inches>
 *       [--color <theme|rgb>] [--fill <theme|rgb>] [--font-size <pt>]
 *       [--bold] [--align <left|center|right>]
 *
 * Theme colors: dk1, dk2, lt1, lt2, accent1-6, hlink, folHlink
 * RGB colors: FF0000, 00FF00, etc.
 */

const fs = require('fs');
const path = require('path');

const EMU_PER_INCH = 914400;
const EMU_PER_PT = 12700;

// Theme color names
const THEME_COLORS = ['dk1', 'dk2', 'lt1', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/**
 * Convert inches to EMU
 */
function inchesToEmu(inches) {
  return Math.round(inches * EMU_PER_INCH);
}

/**
 * Convert points to EMU
 */
function ptToEmu(pt) {
  return Math.round(pt * EMU_PER_PT);
}

/**
 * Check if color is a theme color
 */
function isThemeColor(color) {
  return THEME_COLORS.includes(color);
}

/**
 * Generate color XML
 */
function generateColorXml(color, elementName = 'a:solidFill') {
  if (!color) return '';

  if (isThemeColor(color)) {
    return `<${elementName}><a:schemeClr val="${color}"/></${elementName}>`;
  } else {
    // Assume RGB
    return `<${elementName}><a:srgbClr val="${color.replace('#', '')}"/></${elementName}>`;
  }
}

/**
 * Generate text box shape XML
 */
function generateTextBoxXml(id, options) {
  const {
    x, y, width, height,
    text = '',
    color = 'dk1',  // Default to theme dark color
    fill = null,
    fontSize = 18,
    bold = false,
    italic = false,
    align = 'left',
  } = options;

  const alignMap = { left: 'l', center: 'ctr', right: 'r' };
  const algn = alignMap[align] || 'l';

  const fillXml = fill ? generateColorXml(fill, 'a:solidFill') : '<a:noFill/>';
  const colorXml = generateColorXml(color, 'a:solidFill');

  const boldAttr = bold ? ' b="1"' : '';
  const italicAttr = italic ? ' i="1"' : '';

  return `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${id}" name="TextBox ${id}"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${x}" y="${y}"/>
            <a:ext cx="${width}" cy="${height}"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          ${fillXml}
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="${algn}"/>
            <a:r>
              <a:rPr lang="en-US" sz="${fontSize * 100}"${boldAttr}${italicAttr} dirty="0">
                ${colorXml}
              </a:rPr>
              <a:t>${escapeXml(text)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`;
}

/**
 * Generate rectangle shape XML
 */
function generateRectXml(id, options) {
  const {
    x, y, width, height,
    fill = 'accent1',
    stroke = null,
    strokeWidth = 1,
  } = options;

  const fillXml = fill ? generateColorXml(fill, 'a:solidFill') : '<a:noFill/>';
  const strokeXml = stroke
    ? `<a:ln w="${ptToEmu(strokeWidth)}">${generateColorXml(stroke, 'a:solidFill')}</a:ln>`
    : '<a:ln><a:noFill/></a:ln>';

  return `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${id}" name="Rectangle ${id}"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${x}" y="${y}"/>
            <a:ext cx="${width}" cy="${height}"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          ${fillXml}
          ${strokeXml}
        </p:spPr>
      </p:sp>`;
}

/**
 * Escape XML special characters
 */
function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Get next available shape ID in slide
 */
function getNextShapeId(slideXml) {
  const ids = [];
  const pattern = /<p:cNvPr[^>]*id="(\d+)"/g;
  let match;
  while ((match = pattern.exec(slideXml)) !== null) {
    ids.push(parseInt(match[1], 10));
  }
  return ids.length > 0 ? Math.max(...ids) + 1 : 2;
}

function addShape(workDir, slideNum, shapeType, options) {
  const absWork = path.resolve(workDir);
  const slidePath = path.join(absWork, 'ppt', 'slides', `slide${slideNum}.xml`);

  if (!fs.existsSync(slidePath)) {
    throw new Error(`Slide not found: ${slidePath}`);
  }

  let slideXml = fs.readFileSync(slidePath, 'utf-8');
  const shapeId = getNextShapeId(slideXml);

  // Convert inches to EMU
  const emuOptions = {
    ...options,
    x: inchesToEmu(options.x || 1),
    y: inchesToEmu(options.y || 1),
    width: inchesToEmu(options.width || 4),
    height: inchesToEmu(options.height || 1),
  };

  let shapeXml;
  switch (shapeType) {
    case 'textbox':
    case 'text':
      shapeXml = generateTextBoxXml(shapeId, emuOptions);
      break;
    case 'rect':
    case 'rectangle':
      shapeXml = generateRectXml(shapeId, emuOptions);
      break;
    default:
      throw new Error(`Unknown shape type: ${shapeType}. Use: textbox, rect`);
  }

  // Insert before </p:spTree>
  slideXml = slideXml.replace('</p:spTree>', shapeXml + '\n    </p:spTree>');

  fs.writeFileSync(slidePath, slideXml);
  console.log(`Added ${shapeType} (id=${shapeId}) to slide ${slideNum}`);

  return shapeId;
}

// CLI argument parser
function parseArgs(args) {
  const options = {};
  let workDir = null;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    const next = args[i + 1];

    if (arg === '--slide' && next) { options.slide = parseInt(next, 10); i++; }
    else if (arg === '--type' && next) { options.type = next; i++; }
    else if (arg === '--text' && next) { options.text = next; i++; }
    else if (arg === '--x' && next) { options.x = parseFloat(next); i++; }
    else if (arg === '--y' && next) { options.y = parseFloat(next); i++; }
    else if (arg === '--width' && next) { options.width = parseFloat(next); i++; }
    else if (arg === '--height' && next) { options.height = parseFloat(next); i++; }
    else if (arg === '--color' && next) { options.color = next; i++; }
    else if (arg === '--fill' && next) { options.fill = next; i++; }
    else if (arg === '--stroke' && next) { options.stroke = next; i++; }
    else if (arg === '--font-size' && next) { options.fontSize = parseInt(next, 10); i++; }
    else if (arg === '--bold') { options.bold = true; }
    else if (arg === '--italic') { options.italic = true; }
    else if (arg === '--align' && next) { options.align = next; i++; }
    else if (!arg.startsWith('--')) { workDir = arg; }
  }

  return { workDir, ...options };
}

// CLI execution
if (require.main === module) {
  const options = parseArgs(process.argv.slice(2));

  if (!options.workDir || !options.slide || !options.type) {
    console.error(`Usage: node add-shape.js <work-dir> --slide <n> --type <textbox|rect> [options]

Options:
  --text <string>     Text content (for textbox)
  --x <inches>        X position (default: 1)
  --y <inches>        Y position (default: 1)
  --width <inches>    Width (default: 4)
  --height <inches>   Height (default: 1)
  --color <color>     Text color: theme (dk1, accent1, etc.) or RGB (FF0000)
  --fill <color>      Fill color: theme or RGB
  --stroke <color>    Stroke color (for rect)
  --font-size <pt>    Font size in points (default: 18)
  --bold              Bold text
  --italic            Italic text
  --align <align>     Text alignment: left, center, right

Theme colors: dk1, dk2, lt1, lt2, accent1-6, hlink, folHlink`);
    process.exit(1);
  }

  try {
    addShape(options.workDir, options.slide, options.type, options);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

module.exports = { addShape, THEME_COLORS };
