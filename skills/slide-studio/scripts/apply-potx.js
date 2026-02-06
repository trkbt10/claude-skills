#!/usr/bin/env node
/**
 * Applies a POTX template to the presentation
 * Copies theme, slide masters, and slide layouts from the template
 *
 * Usage: node apply-potx.js <working-directory> <template.potx>
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { parseRelationships, generateRelationships } = require('./lib/relationships');
const { parseContentTypes, generateContentTypes } = require('./lib/content-types');
const { REL_TYPES, CONTENT_TYPES } = require('./lib/constants');

/**
 * Extract POTX to temp directory
 */
async function extractPotx(potxPath) {
  const data = fs.readFileSync(potxPath);
  const zip = await JSZip.loadAsync(data);

  const tempDir = path.join(require('os').tmpdir(), `potx-${Date.now()}`);
  fs.mkdirSync(tempDir, { recursive: true });

  const files = Object.keys(zip.files);
  for (const filePath of files) {
    const zipEntry = zip.files[filePath];
    if (zipEntry.dir) {
      fs.mkdirSync(path.join(tempDir, filePath), { recursive: true });
    } else {
      const content = await zipEntry.async('nodebuffer');
      const fullPath = path.join(tempDir, filePath);
      fs.mkdirSync(path.dirname(fullPath), { recursive: true });
      fs.writeFileSync(fullPath, content);
    }
  }

  return tempDir;
}

/**
 * Copy directory recursively
 */
function copyDir(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  const entries = fs.readdirSync(src, { withFileTypes: true });

  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);

    if (entry.isDirectory()) {
      copyDir(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

/**
 * Remove directory recursively
 */
function removeDir(dir) {
  if (fs.existsSync(dir)) {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        removeDir(fullPath);
      } else {
        fs.unlinkSync(fullPath);
      }
    }
    fs.rmdirSync(dir);
  }
}

/**
 * Update slide relationships to use new layouts
 */
function updateSlideLayoutRefs(workDir, layoutMapping) {
  const slidesDir = path.join(workDir, 'ppt', 'slides');
  const slideRelsDir = path.join(slidesDir, '_rels');

  if (!fs.existsSync(slideRelsDir)) return;

  const relsFiles = fs.readdirSync(slideRelsDir).filter(f => f.endsWith('.rels'));

  for (const relsFile of relsFiles) {
    const relsPath = path.join(slideRelsDir, relsFile);
    let relsXml = fs.readFileSync(relsPath, 'utf-8');

    // Update layout references - keep the same layout number if it exists
    // This preserves the relationship between slides and their intended layouts
    const layoutPattern = /Target="\.\.\/slideLayouts\/slideLayout(\d+)\.xml"/g;
    relsXml = relsXml.replace(layoutPattern, (match, num) => {
      const newNum = layoutMapping[num] || num;
      return `Target="../slideLayouts/slideLayout${newNum}.xml"`;
    });

    fs.writeFileSync(relsPath, relsXml);
  }
}

async function applyPotx(workDir, potxPath) {
  const absWork = path.resolve(workDir);
  const absPotx = path.resolve(potxPath);

  if (!fs.existsSync(absPotx)) {
    throw new Error(`Template file not found: ${absPotx}`);
  }

  console.log('Extracting template...');
  const tempDir = await extractPotx(absPotx);

  try {
    // Directories to replace
    const themeDir = path.join(absWork, 'ppt', 'theme');
    const mastersDir = path.join(absWork, 'ppt', 'slideMasters');
    const layoutsDir = path.join(absWork, 'ppt', 'slideLayouts');

    const tempThemeDir = path.join(tempDir, 'ppt', 'theme');
    const tempMastersDir = path.join(tempDir, 'ppt', 'slideMasters');
    const tempLayoutsDir = path.join(tempDir, 'ppt', 'slideLayouts');

    // Create layout mapping (old number -> new number)
    const layoutMapping = {};

    // Count existing and new layouts
    const existingLayouts = fs.existsSync(layoutsDir)
      ? fs.readdirSync(layoutsDir).filter(f => f.match(/^slideLayout\d+\.xml$/))
      : [];
    const newLayouts = fs.existsSync(tempLayoutsDir)
      ? fs.readdirSync(tempLayoutsDir).filter(f => f.match(/^slideLayout\d+\.xml$/))
      : [];

    // Map existing layout numbers to themselves (if new template has them)
    for (const file of existingLayouts) {
      const num = file.match(/\d+/)[0];
      if (newLayouts.some(f => f.match(/\d+/)[0] === num)) {
        layoutMapping[num] = num;
      } else {
        // Fallback to layout 1 if the layout doesn't exist in new template
        layoutMapping[num] = '1';
      }
    }

    // Remove old theme, masters, layouts
    console.log('Replacing theme and layouts...');
    if (fs.existsSync(themeDir)) removeDir(themeDir);
    if (fs.existsSync(mastersDir)) removeDir(mastersDir);
    if (fs.existsSync(layoutsDir)) removeDir(layoutsDir);

    // Copy new theme, masters, layouts
    if (fs.existsSync(tempThemeDir)) {
      copyDir(tempThemeDir, themeDir);
    }
    if (fs.existsSync(tempMastersDir)) {
      copyDir(tempMastersDir, mastersDir);
    }
    if (fs.existsSync(tempLayoutsDir)) {
      copyDir(tempLayoutsDir, layoutsDir);
    }

    // Update [Content_Types].xml
    console.log('Updating content types...');
    const ctPath = path.join(absWork, '[Content_Types].xml');
    const ct = parseContentTypes(fs.readFileSync(ctPath, 'utf-8'));

    // Remove old layout/master/theme content types
    const partsToRemove = Object.keys(ct.overrides).filter(p =>
      p.includes('/slideLayouts/') ||
      p.includes('/slideMasters/') ||
      p.includes('/theme/')
    );
    for (const part of partsToRemove) {
      delete ct.overrides[part];
    }

    // Add new content types from template
    if (fs.existsSync(tempLayoutsDir)) {
      const layouts = fs.readdirSync(tempLayoutsDir).filter(f => f.endsWith('.xml') && !f.startsWith('_'));
      for (const file of layouts) {
        ct.overrides[`/ppt/slideLayouts/${file}`] = CONTENT_TYPES.slideLayout;
      }
    }

    if (fs.existsSync(tempMastersDir)) {
      const masters = fs.readdirSync(tempMastersDir).filter(f => f.endsWith('.xml') && !f.startsWith('_'));
      for (const file of masters) {
        ct.overrides[`/ppt/slideMasters/${file}`] = CONTENT_TYPES.slideMaster;
      }
    }

    if (fs.existsSync(tempThemeDir)) {
      const themes = fs.readdirSync(tempThemeDir).filter(f => f.endsWith('.xml'));
      for (const file of themes) {
        ct.overrides[`/ppt/theme/${file}`] = CONTENT_TYPES.theme;
      }
    }

    fs.writeFileSync(ctPath, generateContentTypes(ct));

    // Update presentation.xml.rels for masters/theme
    console.log('Updating relationships...');
    const presRelsPath = path.join(absWork, 'ppt', '_rels', 'presentation.xml.rels');
    const presRels = parseRelationships(fs.readFileSync(presRelsPath, 'utf-8'));

    // Remove old master/theme relationships
    const relsToRemove = presRels.filter(r =>
      r.type === REL_TYPES.slideMaster || r.type === REL_TYPES.theme
    );
    for (const rel of relsToRemove) {
      const idx = presRels.indexOf(rel);
      if (idx >= 0) presRels.splice(idx, 1);
    }

    // Add new master/theme relationships from template
    const tempPresRelsPath = path.join(tempDir, 'ppt', '_rels', 'presentation.xml.rels');
    if (fs.existsSync(tempPresRelsPath)) {
      const tempRels = parseRelationships(fs.readFileSync(tempPresRelsPath, 'utf-8'));
      const maxRId = Math.max(...presRels.map(r => parseInt(r.id.replace('rId', ''), 10)));
      let nextRId = maxRId + 1;

      for (const rel of tempRels) {
        if (rel.type === REL_TYPES.slideMaster || rel.type === REL_TYPES.theme) {
          presRels.push({
            id: `rId${nextRId++}`,
            type: rel.type,
            target: rel.target,
          });
        }
      }
    }

    fs.writeFileSync(presRelsPath, generateRelationships(presRels));

    // Update presentation.xml to reference correct slide master
    const presPath = path.join(absWork, 'ppt', 'presentation.xml');
    let presXml = fs.readFileSync(presPath, 'utf-8');

    // Find the new slide master rId
    const masterRel = presRels.find(r => r.type === REL_TYPES.slideMaster);
    if (masterRel) {
      // Update sldMasterIdLst to use the correct rId
      presXml = presXml.replace(
        /<p:sldMasterId[^>]*r:id="[^"]*"[^>]*\/>/g,
        `<p:sldMasterId id="2147483648" r:id="${masterRel.id}"/>`
      );
      fs.writeFileSync(presPath, presXml);
    }

    // Update slide layout references
    updateSlideLayoutRefs(absWork, layoutMapping);

    console.log('Template applied successfully!');
  } finally {
    // Clean up temp directory
    removeDir(tempDir);
  }
}

// CLI execution
if (require.main === module) {
  const [, , workDir, potxPath] = process.argv;

  if (!workDir || !potxPath) {
    console.error('Usage: node apply-potx.js <working-directory> <template.potx>');
    process.exit(1);
  }

  applyPotx(workDir, potxPath)
    .then(() => process.exit(0))
    .catch((err) => {
      console.error('Error:', err.message);
      process.exit(1);
    });
}

module.exports = { applyPotx };
