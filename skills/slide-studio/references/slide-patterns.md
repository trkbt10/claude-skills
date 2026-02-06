# Slide XML Patterns

## Title Slide (Layout 1)

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title 1"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US"/><a:t>Presentation Title</a:t></a:r></a:p>
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
    <a:p><a:r><a:rPr lang="en-US"/><a:t>Subtitle text</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
```

## Title and Content (Layout 2)

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title 1"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:t>Slide Title</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="3" name="Content 2"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p>
      <a:pPr lvl="0"/>
      <a:r><a:t>First bullet point</a:t></a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="0"/>
      <a:r><a:t>Second bullet point</a:t></a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="1"/>
      <a:r><a:t>Indented bullet</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

## Multiline Text with Formatting

```xml
<p:txBody>
  <a:bodyPr/>
  <a:lstStyle/>
  <a:p>
    <a:r>
      <a:rPr lang="en-US" b="1"/>
      <a:t>Bold text</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:r>
      <a:rPr lang="en-US" i="1"/>
      <a:t>Italic text</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:r>
      <a:rPr lang="en-US"/>
      <a:t>Normal </a:t>
    </a:r>
    <a:r>
      <a:rPr lang="en-US" b="1"/>
      <a:t>bold</a:t>
    </a:r>
    <a:r>
      <a:rPr lang="en-US"/>
      <a:t> mixed</a:t>
    </a:r>
  </a:p>
</p:txBody>
```

## Table

```xml
<a:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="4" name="Table 1"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="457200" y="1600200"/>
    <a:ext cx="8229600" cy="2000000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr firstRow="1"/>
        <a:tblGrid>
          <a:gridCol w="2743200"/>
          <a:gridCol w="2743200"/>
          <a:gridCol w="2743200"/>
        </a:tblGrid>
        <a:tr h="370840">
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>Header 1</a:t></a:r></a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>Header 2</a:t></a:r></a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>Header 3</a:t></a:r></a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>
        </a:tr>
        <a:tr h="370840">
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>Cell 1</a:t></a:r></a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>Cell 2</a:t></a:r></a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>Cell 3</a:t></a:r></a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</a:graphicFrame>
```

## Shape (Rectangle with fill)

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="5" name="Rectangle 1"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="1000000" y="1000000"/>
      <a:ext cx="2000000" cy="1000000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill>
      <a:srgbClr val="4472C4"/>
    </a:solidFill>
    <a:ln w="12700">
      <a:solidFill>
        <a:srgbClr val="2F528F"/>
      </a:solidFill>
    </a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        </a:rPr>
        <a:t>Shape Text</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

## Common Preset Geometries

| prst | Shape |
|------|-------|
| rect | Rectangle |
| roundRect | Rounded Rectangle |
| ellipse | Ellipse |
| triangle | Triangle |
| rtTriangle | Right Triangle |
| parallelogram | Parallelogram |
| trapezoid | Trapezoid |
| pentagon | Pentagon |
| hexagon | Hexagon |
| star5 | 5-Point Star |
| rightArrow | Right Arrow |
| leftArrow | Left Arrow |
| upArrow | Up Arrow |
| downArrow | Down Arrow |
| cloud | Cloud |
| heart | Heart |
