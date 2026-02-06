# OOXML/PresentationML Reference

## PPTX Structure

PPTX is a ZIP archive containing XML files:

```
presentation.pptx/
├── [Content_Types].xml       # Content type definitions
├── _rels/.rels               # Root relationships
├── ppt/
│   ├── presentation.xml      # Main presentation
│   ├── _rels/presentation.xml.rels
│   ├── slides/
│   │   ├── slide1.xml
│   │   └── _rels/slide1.xml.rels
│   ├── slideLayouts/
│   ├── slideMasters/
│   └── theme/
└── docProps/
    ├── core.xml
    └── app.xml
```

## XML Namespaces

```
p:  http://schemas.openxmlformats.org/presentationml/2006/main
a:  http://schemas.openxmlformats.org/drawingml/2006/main
r:  http://schemas.openxmlformats.org/officeDocument/2006/relationships
```

## EMU (English Metric Units)

- 914400 EMU = 1 inch
- 360000 EMU = 1 cm
- 12700 EMU = 1 point
- 9525 EMU = 1 pixel (at 96 DPI)

Standard slide size (4:3): 9144000 x 6858000 EMU (10" x 7.5")

## Relationship Types

| Type | URI |
|------|-----|
| slide | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide` |
| slideLayout | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout` |
| slideMaster | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster` |
| theme | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme` |
| image | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image` |

## Content Types

| Part | Content Type |
|------|--------------|
| presentation | `application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml` |
| slide | `application/vnd.openxmlformats-officedocument.presentationml.slide+xml` |
| slideLayout | `application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml` |
| slideMaster | `application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml` |
| theme | `application/vnd.openxmlformats-officedocument.theme+xml` |

## Placeholder Types

| Type | Description |
|------|-------------|
| title | Standard title |
| ctrTitle | Centered title (title slide) |
| subTitle | Subtitle |
| body | Body text |
| dt | Date/time |
| ftr | Footer |
| sldNum | Slide number |
| pic | Picture placeholder |
| tbl | Table placeholder |
| chart | Chart placeholder |

## Adding a Slide

1. Create `ppt/slides/slideN.xml`
2. Create `ppt/slides/_rels/slideN.xml.rels` (link to layout)
3. Add relationship in `ppt/_rels/presentation.xml.rels`
4. Add `<p:sldId>` to `ppt/presentation.xml`
5. Add Override in `[Content_Types].xml`

## Shape Structure

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title 1"/>
    <p:cNvSpPr/>
    <p:nvPr>
      <p:ph type="title"/>
    </p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="457200" y="274638"/>
      <a:ext cx="8229600" cy="1143000"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US"/>
        <a:t>Text content</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

## Image Structure

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="4" name="Picture 1"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="1000000" y="1000000"/>
      <a:ext cx="3000000" cy="2000000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

## Text Formatting

```xml
<!-- Bold -->
<a:rPr b="1"/>

<!-- Italic -->
<a:rPr i="1"/>

<!-- Underline -->
<a:rPr u="sng"/>

<!-- Font size (in hundredths of a point) -->
<a:rPr sz="2400"/> <!-- 24pt -->

<!-- Font color -->
<a:rPr>
  <a:solidFill>
    <a:srgbClr val="FF0000"/>
  </a:solidFill>
</a:rPr>
```

## Bullet Lists

```xml
<a:p>
  <a:pPr lvl="0">
    <a:buChar char="•"/>
  </a:pPr>
  <a:r><a:t>First item</a:t></a:r>
</a:p>
<a:p>
  <a:pPr lvl="1">
    <a:buChar char="–"/>
  </a:pPr>
  <a:r><a:t>Sub-item</a:t></a:r>
</a:p>
```

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| "Unreadable content" | Missing Content_Types entry | Add Override for the part |
| Slide not showing | Missing sldId in presentation.xml | Add sldId with correct r:id |
| Image not displaying | Missing relationship | Add relationship in slide .rels |
| Layout not applied | Wrong slideLayout reference | Fix Target in slide .rels |
