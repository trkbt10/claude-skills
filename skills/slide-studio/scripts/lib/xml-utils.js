/**
 * XML Utilities for PPTX manipulation
 * Uses regex-based parsing for known PPTX patterns
 */

/**
 * Escape special XML characters
 */
function escapeXml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Unescape XML entities
 */
function unescapeXml(str) {
  if (!str) return '';
  return str
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

/**
 * Get attribute value from XML element
 */
function getAttribute(xml, attrName) {
  const pattern = new RegExp(`${attrName}\\s*=\\s*["']([^"']*)["']`);
  const match = xml.match(pattern);
  return match ? match[1] : null;
}

/**
 * Set attribute value in XML element
 */
function setAttribute(xml, attrName, value) {
  const pattern = new RegExp(`(${attrName}\\s*=\\s*["'])([^"']*)(["'])`);
  if (pattern.test(xml)) {
    return xml.replace(pattern, `$1${escapeXml(value)}$3`);
  }
  // Add attribute if not exists (before closing > or />)
  return xml.replace(/(\/?>)/, ` ${attrName}="${escapeXml(value)}"$1`);
}

/**
 * Find all elements matching a tag pattern
 */
function findElements(xml, tagName) {
  const pattern = new RegExp(`<${tagName}[^>]*(?:/>|>[\\s\\S]*?</${tagName}>)`, 'g');
  return xml.match(pattern) || [];
}

/**
 * Find element by tag and attribute
 */
function findElementByAttr(xml, tagName, attrName, attrValue) {
  const elements = findElements(xml, tagName);
  return elements.find(el => getAttribute(el, attrName) === attrValue) || null;
}

/**
 * Get inner content of an element
 */
function getInnerContent(xml, tagName) {
  const pattern = new RegExp(`<${tagName}[^>]*>([\\s\\S]*?)</${tagName}>`);
  const match = xml.match(pattern);
  return match ? match[1] : null;
}

/**
 * Replace inner content of an element
 */
function setInnerContent(xml, tagName, content) {
  const pattern = new RegExp(`(<${tagName}[^>]*>)[\\s\\S]*?(</${tagName}>)`);
  return xml.replace(pattern, `$1${content}$2`);
}

/**
 * Get all text content (a:t elements) from XML
 */
function getAllText(xml) {
  const texts = [];
  const pattern = /<a:t>([^<]*)<\/a:t>/g;
  let match;
  while ((match = pattern.exec(xml)) !== null) {
    texts.push(unescapeXml(match[1]));
  }
  return texts;
}

/**
 * Replace text in a:t element
 */
function replaceText(xml, oldText, newText) {
  const escaped = escapeXml(newText);
  return xml.replace(
    new RegExp(`(<a:t>)${escapeRegex(oldText)}(</a:t>)`, 'g'),
    `$1${escaped}$2`
  );
}

/**
 * Set text for a placeholder type
 */
function setPlaceholderText(xml, phType, text) {
  // Find the shape with the placeholder type
  const phPattern = new RegExp(
    `(<p:sp[^>]*>[\\s\\S]*?<p:ph[^>]*type="${phType}"[^>]*/>[\\s\\S]*?<p:txBody>[\\s\\S]*?)(<a:t>)[^<]*(</a:t>)`,
    'g'
  );
  return xml.replace(phPattern, `$1$2${escapeXml(text)}$3`);
}

/**
 * Find placeholder shape by type
 */
function findPlaceholder(xml, phType) {
  const pattern = new RegExp(
    `<p:sp[^>]*>[\\s\\S]*?<p:ph[^>]*type="${phType}"[^>]*/>[\\s\\S]*?</p:sp>`,
    'g'
  );
  const matches = xml.match(pattern);
  return matches ? matches[0] : null;
}

/**
 * Get next available ID in the document
 */
function getNextId(xml) {
  const ids = [];
  const pattern = /\bid\s*=\s*["'](\d+)["']/g;
  let match;
  while ((match = pattern.exec(xml)) !== null) {
    ids.push(parseInt(match[1], 10));
  }
  return ids.length > 0 ? Math.max(...ids) + 1 : 1;
}

/**
 * Escape string for use in regex
 */
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Generate XML declaration
 */
function xmlDeclaration() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
}

/**
 * Pretty format XML (basic indentation)
 */
function formatXml(xml) {
  let formatted = '';
  let indent = 0;
  const lines = xml.replace(/>\s*</g, '>\n<').split('\n');

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    if (trimmed.startsWith('</')) {
      indent = Math.max(0, indent - 1);
    }

    formatted += '  '.repeat(indent) + trimmed + '\n';

    if (trimmed.startsWith('<') && !trimmed.startsWith('</') &&
        !trimmed.startsWith('<?') && !trimmed.endsWith('/>') &&
        !trimmed.includes('</')) {
      indent++;
    }
  }

  return formatted;
}

module.exports = {
  escapeXml,
  unescapeXml,
  getAttribute,
  setAttribute,
  findElements,
  findElementByAttr,
  getInnerContent,
  setInnerContent,
  getAllText,
  replaceText,
  setPlaceholderText,
  findPlaceholder,
  getNextId,
  escapeRegex,
  xmlDeclaration,
  formatXml,
};
