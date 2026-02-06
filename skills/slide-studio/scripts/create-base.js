#!/usr/bin/env node
/**
 * Creates a minimal base PPTX with standard layouts
 * Usage: node create-base.js <output.pptx>
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const LAYOUTS = [
  { name: 'Title Slide', placeholders: ['ctrTitle', 'subTitle'] },
  { name: 'Title and Content', placeholders: ['title', 'body'] },
  { name: 'Section Header', placeholders: ['title', 'subTitle'] },
  { name: 'Two Content', placeholders: ['title', 'body', 'body'] },
  { name: 'Comparison', placeholders: ['title', 'body', 'body'] },
  { name: 'Title Only', placeholders: ['title'] },
  { name: 'Blank', placeholders: [] },
  { name: 'Content with Caption', placeholders: ['title', 'body'] },
  { name: 'Picture with Caption', placeholders: ['title', 'body'] },
  { name: 'Title and Vertical Text', placeholders: ['title', 'body'] },
  { name: 'Vertical Title and Text', placeholders: ['title', 'body'] },
];

function xmlDeclaration() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
}

function createContentTypes() {
  let xml = xmlDeclaration();
  xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';
  xml += '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n';
  xml += '  <Default Extension="xml" ContentType="application/xml"/>\n';
  xml += '  <Default Extension="png" ContentType="image/png"/>\n';
  xml += '  <Default Extension="jpeg" ContentType="image/jpeg"/>\n';
  xml += '  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>\n';
  xml += '  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>\n';
  xml += '  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>\n';
  xml += '  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>\n';
  xml += '  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>\n';

  for (let i = 1; i <= LAYOUTS.length; i++) {
    xml += `  <Override PartName="/ppt/slideLayouts/slideLayout${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>\n`;
  }

  xml += '  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\n';
  xml += '  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n';
  xml += '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n';
  xml += '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n';
  xml += '</Types>';
  return xml;
}

function createRootRels() {
  return xmlDeclaration() +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function createPresentation() {
  return xmlDeclaration() +
`<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function createPresentationRels() {
  return xmlDeclaration() +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;
}

function createPresProps() {
  return xmlDeclaration() +
`<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
}

function createViewProps() {
  return xmlDeclaration() +
`<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>
</p:viewPr>`;
}

function createTableStyles() {
  return xmlDeclaration() +
`<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}

function createTheme() {
  return xmlDeclaration() +
`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

function createSlideMaster() {
  let layoutRefs = '';
  for (let i = 1; i <= LAYOUTS.length; i++) {
    layoutRefs += `    <p:sldLayoutId id="${2147483649 + i}" r:id="rId${i}"/>\n`;
  }

  return xmlDeclaration() +
`<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
${layoutRefs}  </p:sldLayoutIdLst>
</p:sldMaster>`;
}

function createSlideMasterRels() {
  let rels = '';
  for (let i = 1; i <= LAYOUTS.length; i++) {
    rels += `  <Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${i}.xml"/>\n`;
  }
  rels += `  <Relationship Id="rId${LAYOUTS.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>`;

  return xmlDeclaration() +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
</Relationships>`;
}

function createSlideLayout(index, layout) {
  const { name, placeholders } = layout;
  let shapes = '';
  let id = 2;

  for (const ph of placeholders) {
    const isTitle = ph === 'title' || ph === 'ctrTitle';
    const y = isTitle ? 274638 : 1600200;
    const h = isTitle ? 1143000 : 4525963;

    shapes += `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${id}" name="${ph} ${id}"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="${ph}"${id > 2 ? ` idx="${id - 2}"` : ''}/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="838200" y="${y}"/><a:ext cx="10515600" cy="${h}"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>`;
    id++;
  }

  return xmlDeclaration() +
`<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj">
  <p:cSld name="${name}">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
}

function createSlideLayoutRels() {
  return xmlDeclaration() +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}

function createSlide() {
  return xmlDeclaration() +
`<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US"/><a:t>Title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Subtitle 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US"/><a:t>Subtitle</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function createSlideRels() {
  return xmlDeclaration() +
`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
}

function createCoreProps() {
  return xmlDeclaration() +
`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Presentation</dc:title>
  <dc:creator>PPTX Builder</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`;
}

function createAppProps() {
  return xmlDeclaration() +
`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>PPTX Builder</Application>
  <Slides>1</Slides>
</Properties>`;
}

async function createBase(outputPath) {
  const zip = new JSZip();

  // Root files
  zip.file('[Content_Types].xml', createContentTypes());
  zip.file('_rels/.rels', createRootRels());

  // docProps
  zip.file('docProps/core.xml', createCoreProps());
  zip.file('docProps/app.xml', createAppProps());

  // ppt root
  zip.file('ppt/presentation.xml', createPresentation());
  zip.file('ppt/_rels/presentation.xml.rels', createPresentationRels());
  zip.file('ppt/presProps.xml', createPresProps());
  zip.file('ppt/viewProps.xml', createViewProps());
  zip.file('ppt/tableStyles.xml', createTableStyles());

  // Theme
  zip.file('ppt/theme/theme1.xml', createTheme());

  // Slide Master
  zip.file('ppt/slideMasters/slideMaster1.xml', createSlideMaster());
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', createSlideMasterRels());

  // Slide Layouts
  for (let i = 0; i < LAYOUTS.length; i++) {
    zip.file(`ppt/slideLayouts/slideLayout${i + 1}.xml`, createSlideLayout(i + 1, LAYOUTS[i]));
    zip.file(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, createSlideLayoutRels());
  }

  // Initial slide
  zip.file('ppt/slides/slide1.xml', createSlide());
  zip.file('ppt/slides/_rels/slide1.xml.rels', createSlideRels());

  // Generate ZIP
  const content = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  fs.writeFileSync(outputPath, content);
  console.log(`Created: ${outputPath}`);
  console.log(`Layouts: ${LAYOUTS.length}`);
}

// CLI execution
if (require.main === module) {
  const outputPath = process.argv[2] || 'base.pptx';

  createBase(outputPath)
    .then(() => process.exit(0))
    .catch((err) => {
      console.error('Error:', err.message);
      process.exit(1);
    });
}

module.exports = { createBase };
