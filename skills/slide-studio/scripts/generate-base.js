#!/usr/bin/env node
/**
 * Generates a minimal base.pptx with 11 standard layouts
 * Usage: node generate-base.js <output-path>
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// Standard namespaces
const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
};

// Layout definitions with placeholder configurations
const LAYOUTS = [
  { name: 'Title Slide', type: 'title', placeholders: [
    { type: 'ctrTitle', x: 685800, y: 2130425, cx: 7772400, cy: 1470025 },
    { type: 'subTitle', x: 1371600, y: 3886200, cx: 6400800, cy: 1752600, idx: 1 },
  ]},
  { name: 'Title and Content', type: 'obj', placeholders: [
    { type: 'title', x: 457200, y: 274638, cx: 8229600, cy: 1143000 },
    { type: 'body', x: 457200, y: 1600200, cx: 8229600, cy: 4525963, idx: 1 },
  ]},
  { name: 'Section Header', type: 'secHead', placeholders: [
    { type: 'title', x: 722313, y: 4406900, cx: 7772400, cy: 1362075 },
    { type: 'body', x: 722313, y: 2906713, cx: 7772400, cy: 1500187, idx: 1 },
  ]},
  { name: 'Two Content', type: 'twoObj', placeholders: [
    { type: 'title', x: 457200, y: 274638, cx: 8229600, cy: 1143000 },
    { type: 'body', x: 457200, y: 1600200, cx: 4038600, cy: 4525963, idx: 1 },
    { type: 'body', x: 4648200, y: 1600200, cx: 4038600, cy: 4525963, idx: 2 },
  ]},
  { name: 'Comparison', type: 'twoTxTwoObj', placeholders: [
    { type: 'title', x: 457200, y: 274638, cx: 8229600, cy: 1143000 },
    { type: 'body', x: 457200, y: 1535113, cx: 4040188, cy: 639762, idx: 1 },
    { type: 'body', x: 457200, y: 2174875, cx: 4040188, cy: 3951288, idx: 2 },
    { type: 'body', x: 4645025, y: 1535113, cx: 4041775, cy: 639762, idx: 3 },
    { type: 'body', x: 4645025, y: 2174875, cx: 4041775, cy: 3951288, idx: 4 },
  ]},
  { name: 'Title Only', type: 'titleOnly', placeholders: [
    { type: 'title', x: 457200, y: 274638, cx: 8229600, cy: 1143000 },
  ]},
  { name: 'Blank', type: 'blank', placeholders: []},
  { name: 'Content with Caption', type: 'objTx', placeholders: [
    { type: 'title', x: 457200, y: 273050, cx: 3008313, cy: 1162050 },
    { type: 'body', x: 3575050, y: 273050, cx: 5111750, cy: 5853113, idx: 1 },
    { type: 'body', x: 457200, y: 1435100, cx: 3008313, cy: 4691063, idx: 2 },
  ]},
  { name: 'Picture with Caption', type: 'picTx', placeholders: [
    { type: 'title', x: 1792288, y: 4800600, cx: 5486400, cy: 566738 },
    { type: 'body', x: 1792288, y: 5367338, cx: 5486400, cy: 804862, idx: 1 },
    { type: 'pic', x: 1792288, y: 612775, cx: 5486400, cy: 4114800, idx: 2 },
  ]},
  { name: 'Title and Vertical Text', type: 'vertTx', placeholders: [
    { type: 'title', x: 457200, y: 274638, cx: 8229600, cy: 1143000 },
    { type: 'body', x: 457200, y: 1600200, cx: 8229600, cy: 4525963, idx: 1, vert: true },
  ]},
  { name: 'Vertical Title and Text', type: 'vertTitleAndTx', placeholders: [
    { type: 'title', x: 6629400, y: 274638, cx: 2057400, cy: 5851525, vert: true },
    { type: 'body', x: 457200, y: 274638, cx: 6019800, cy: 5851525, idx: 1, vert: true },
  ]},
];

function generateContentTypes() {
  let xml = XML_DECLARATION;
  xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';
  xml += '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n';
  xml += '  <Default Extension="xml" ContentType="application/xml"/>\n';
  xml += '  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>\n';
  xml += '  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>\n';
  xml += '  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>\n';
  xml += '  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>\n';
  xml += '  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>\n';

  for (let i = 1; i <= 11; i++) {
    xml += `  <Override PartName="/ppt/slideLayouts/slideLayout${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>\n`;
  }

  xml += '  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\n';
  xml += '  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n';
  xml += '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n';
  xml += '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n';
  xml += '</Types>';
  return xml;
}

function generateRootRels() {
  let xml = XML_DECLARATION;
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
  xml += '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>\n';
  xml += '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\n';
  xml += '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\n';
  xml += '</Relationships>';
  return xml;
}

function generatePresentation() {
  let xml = XML_DECLARATION;
  xml += `<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" saveSubsetFonts="1">\n`;
  xml += '  <p:sldMasterIdLst>\n';
  xml += '    <p:sldMasterId id="2147483648" r:id="rId1"/>\n';
  xml += '  </p:sldMasterIdLst>\n';
  xml += '  <p:sldIdLst>\n';
  xml += '    <p:sldId id="256" r:id="rId2"/>\n';
  xml += '  </p:sldIdLst>\n';
  xml += '  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>\n';
  xml += '  <p:notesSz cx="6858000" cy="9144000"/>\n';
  xml += '</p:presentation>';
  return xml;
}

function generatePresentationRels() {
  let xml = XML_DECLARATION;
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
  xml += '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>\n';
  xml += '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>\n';
  xml += '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>\n';
  xml += '  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>\n';
  xml += '  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>\n';
  xml += '  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>\n';
  xml += '</Relationships>';
  return xml;
}

function generatePresProps() {
  return XML_DECLARATION + '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>';
}

function generateViewProps() {
  return XML_DECLARATION + '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="68" d="100"/><a:sy n="68" d="100"/></p:scale><p:origin x="-1392" y="-96"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="2880"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>';
}

function generateTableStyles() {
  return XML_DECLARATION + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
}

function generateTheme() {
  let xml = XML_DECLARATION;
  xml += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">\n';
  xml += '  <a:themeElements>\n';
  xml += '    <a:clrScheme name="Office">\n';
  xml += '      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>\n';
  xml += '      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\n';
  xml += '      <a:dk2><a:srgbClr val="44546A"/></a:dk2>\n';
  xml += '      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>\n';
  xml += '      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>\n';
  xml += '      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>\n';
  xml += '      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>\n';
  xml += '      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>\n';
  xml += '      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>\n';
  xml += '      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>\n';
  xml += '      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>\n';
  xml += '      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>\n';
  xml += '    </a:clrScheme>\n';
  xml += '    <a:fontScheme name="Office">\n';
  xml += '      <a:majorFont>\n';
  xml += '        <a:latin typeface="Calibri Light"/>\n';
  xml += '        <a:ea typeface=""/>\n';
  xml += '        <a:cs typeface=""/>\n';
  xml += '      </a:majorFont>\n';
  xml += '      <a:minorFont>\n';
  xml += '        <a:latin typeface="Calibri"/>\n';
  xml += '        <a:ea typeface=""/>\n';
  xml += '        <a:cs typeface=""/>\n';
  xml += '      </a:minorFont>\n';
  xml += '    </a:fontScheme>\n';
  xml += '    <a:fmtScheme name="Office">\n';
  xml += '      <a:fillStyleLst>\n';
  xml += '        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\n';
  xml += '        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>\n';
  xml += '        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>\n';
  xml += '      </a:fillStyleLst>\n';
  xml += '      <a:lnStyleLst>\n';
  xml += '        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\n';
  xml += '        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\n';
  xml += '        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\n';
  xml += '      </a:lnStyleLst>\n';
  xml += '      <a:effectStyleLst>\n';
  xml += '        <a:effectStyle><a:effectLst/></a:effectStyle>\n';
  xml += '        <a:effectStyle><a:effectLst/></a:effectStyle>\n';
  xml += '        <a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>\n';
  xml += '      </a:effectStyleLst>\n';
  xml += '      <a:bgFillStyleLst>\n';
  xml += '        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\n';
  xml += '        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>\n';
  xml += '        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>\n';
  xml += '      </a:bgFillStyleLst>\n';
  xml += '    </a:fmtScheme>\n';
  xml += '  </a:themeElements>\n';
  xml += '  <a:objectDefaults/>\n';
  xml += '  <a:extraClrSchemeLst/>\n';
  xml += '</a:theme>';
  return xml;
}

function generatePlaceholder(ph, shapeId) {
  const vert = ph.vert ? ' vert="eaVert"' : '';
  let xml = `    <p:sp>\n`;
  xml += `      <p:nvSpPr>\n`;
  xml += `        <p:cNvPr id="${shapeId}" name="${ph.type} ${shapeId}"/>\n`;
  xml += `        <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>\n`;
  xml += `        <p:nvPr><p:ph type="${ph.type}"${ph.idx !== undefined ? ` idx="${ph.idx}"` : ''}/></p:nvPr>\n`;
  xml += `      </p:nvSpPr>\n`;
  xml += `      <p:spPr>\n`;
  xml += `        <a:xfrm>\n`;
  xml += `          <a:off x="${ph.x}" y="${ph.y}"/>\n`;
  xml += `          <a:ext cx="${ph.cx}" cy="${ph.cy}"/>\n`;
  xml += `        </a:xfrm>\n`;
  xml += `        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n`;
  xml += `      </p:spPr>\n`;
  xml += `      <p:txBody>\n`;
  xml += `        <a:bodyPr${vert}/>\n`;
  xml += `        <a:lstStyle/>\n`;
  xml += `        <a:p><a:endParaRPr lang="en-US"/></a:p>\n`;
  xml += `      </p:txBody>\n`;
  xml += `    </p:sp>\n`;
  return xml;
}

function generateSlideMaster() {
  let xml = XML_DECLARATION;
  xml += `<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">\n`;
  xml += '  <p:cSld>\n';
  xml += '    <p:bg>\n';
  xml += '      <p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>\n';
  xml += '    </p:bg>\n';
  xml += '    <p:spTree>\n';
  xml += '      <p:nvGrpSpPr>\n';
  xml += '        <p:cNvPr id="1" name=""/>\n';
  xml += '        <p:cNvGrpSpPr/>\n';
  xml += '        <p:nvPr/>\n';
  xml += '      </p:nvGrpSpPr>\n';
  xml += '      <p:grpSpPr>\n';
  xml += '        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>\n';
  xml += '      </p:grpSpPr>\n';
  xml += '    </p:spTree>\n';
  xml += '  </p:cSld>\n';
  xml += '  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>\n';
  xml += '  <p:sldLayoutIdLst>\n';
  for (let i = 1; i <= 11; i++) {
    xml += `    <p:sldLayoutId id="${2147483649 + i}" r:id="rId${i}"/>\n`;
  }
  xml += '  </p:sldLayoutIdLst>\n';
  xml += '  <p:txStyles>\n';
  xml += '    <p:titleStyle>\n';
  xml += '      <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>\n';
  xml += '    </p:titleStyle>\n';
  xml += '    <p:bodyStyle>\n';
  xml += '      <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>\n';
  xml += '      <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>\n';
  xml += '      <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>\n';
  xml += '      <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>\n';
  xml += '      <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>\n';
  xml += '    </p:bodyStyle>\n';
  xml += '    <p:otherStyle>\n';
  xml += '      <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>\n';
  xml += '    </p:otherStyle>\n';
  xml += '  </p:txStyles>\n';
  xml += '</p:sldMaster>';
  return xml;
}

function generateSlideMasterRels() {
  let xml = XML_DECLARATION;
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
  for (let i = 1; i <= 11; i++) {
    xml += `  <Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${i}.xml"/>\n`;
  }
  xml += '  <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>\n';
  xml += '</Relationships>';
  return xml;
}

function generateSlideLayout(layout, index) {
  let xml = XML_DECLARATION;
  xml += `<p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="${layout.type}" preserve="1">\n`;
  xml += `  <p:cSld name="${layout.name}">\n`;
  xml += '    <p:spTree>\n';
  xml += '      <p:nvGrpSpPr>\n';
  xml += '        <p:cNvPr id="1" name=""/>\n';
  xml += '        <p:cNvGrpSpPr/>\n';
  xml += '        <p:nvPr/>\n';
  xml += '      </p:nvGrpSpPr>\n';
  xml += '      <p:grpSpPr>\n';
  xml += '        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>\n';
  xml += '      </p:grpSpPr>\n';

  let shapeId = 2;
  for (const ph of layout.placeholders) {
    xml += generatePlaceholder(ph, shapeId++);
  }

  xml += '    </p:spTree>\n';
  xml += '  </p:cSld>\n';
  xml += '  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\n';
  xml += '</p:sldLayout>';
  return xml;
}

function generateSlideLayoutRels() {
  let xml = XML_DECLARATION;
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
  xml += '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>\n';
  xml += '</Relationships>';
  return xml;
}

function generateSlide() {
  let xml = XML_DECLARATION;
  xml += `<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">\n`;
  xml += '  <p:cSld>\n';
  xml += '    <p:spTree>\n';
  xml += '      <p:nvGrpSpPr>\n';
  xml += '        <p:cNvPr id="1" name=""/>\n';
  xml += '        <p:cNvGrpSpPr/>\n';
  xml += '        <p:nvPr/>\n';
  xml += '      </p:nvGrpSpPr>\n';
  xml += '      <p:grpSpPr>\n';
  xml += '        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>\n';
  xml += '      </p:grpSpPr>\n';
  xml += '    </p:spTree>\n';
  xml += '  </p:cSld>\n';
  xml += '  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\n';
  xml += '</p:sld>';
  return xml;
}

function generateSlideRels() {
  let xml = XML_DECLARATION;
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
  xml += '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>\n';
  xml += '</Relationships>';
  return xml;
}

function generateCoreProps() {
  const now = new Date().toISOString();
  let xml = XML_DECLARATION;
  xml += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n';
  xml += '  <dc:title>Presentation</dc:title>\n';
  xml += '  <dc:creator>slide-studio</dc:creator>\n';
  xml += `  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>\n`;
  xml += `  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>\n`;
  xml += '</cp:coreProperties>';
  return xml;
}

function generateAppProps() {
  let xml = XML_DECLARATION;
  xml += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n';
  xml += '  <TotalTime>0</TotalTime>\n';
  xml += '  <Words>0</Words>\n';
  xml += '  <Application>slide-studio</Application>\n';
  xml += '  <PresentationFormat>On-screen Show (4:3)</PresentationFormat>\n';
  xml += '  <Paragraphs>0</Paragraphs>\n';
  xml += '  <Slides>1</Slides>\n';
  xml += '  <Notes>0</Notes>\n';
  xml += '  <HiddenSlides>0</HiddenSlides>\n';
  xml += '  <MMClips>0</MMClips>\n';
  xml += '  <ScaleCrop>false</ScaleCrop>\n';
  xml += '  <LinksUpToDate>false</LinksUpToDate>\n';
  xml += '  <SharedDoc>false</SharedDoc>\n';
  xml += '  <HyperlinksChanged>false</HyperlinksChanged>\n';
  xml += '  <AppVersion>16.0000</AppVersion>\n';
  xml += '</Properties>';
  return xml;
}

async function generateBase(outputPath) {
  const zip = new JSZip();

  // Add core files
  zip.file('[Content_Types].xml', generateContentTypes());
  zip.file('_rels/.rels', generateRootRels());
  zip.file('ppt/presentation.xml', generatePresentation());
  zip.file('ppt/_rels/presentation.xml.rels', generatePresentationRels());
  zip.file('ppt/presProps.xml', generatePresProps());
  zip.file('ppt/viewProps.xml', generateViewProps());
  zip.file('ppt/tableStyles.xml', generateTableStyles());
  zip.file('ppt/theme/theme1.xml', generateTheme());
  zip.file('ppt/slideMasters/slideMaster1.xml', generateSlideMaster());
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', generateSlideMasterRels());

  // Add all 11 layouts
  for (let i = 0; i < LAYOUTS.length; i++) {
    zip.file(`ppt/slideLayouts/slideLayout${i + 1}.xml`, generateSlideLayout(LAYOUTS[i], i + 1));
    zip.file(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, generateSlideLayoutRels());
  }

  // Add initial slide
  zip.file('ppt/slides/slide1.xml', generateSlide());
  zip.file('ppt/slides/_rels/slide1.xml.rels', generateSlideRels());

  // Add document properties
  zip.file('docProps/core.xml', generateCoreProps());
  zip.file('docProps/app.xml', generateAppProps());

  // Generate and write
  const content = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  const absOutput = path.resolve(outputPath);
  fs.mkdirSync(path.dirname(absOutput), { recursive: true });
  fs.writeFileSync(absOutput, content);

  console.log(`Generated: ${absOutput}`);
  return absOutput;
}

// CLI execution
if (require.main === module) {
  const outputPath = process.argv[2] || 'assets/base.pptx';

  generateBase(outputPath)
    .then(() => process.exit(0))
    .catch((err) => {
      console.error('Error:', err.message);
      process.exit(1);
    });
}

module.exports = { generateBase };
