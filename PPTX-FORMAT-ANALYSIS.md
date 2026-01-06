# PPTX File Format Analysis

## File Structure

A PPTX file is a ZIP archive containing XML files and binary resources.

### Root Structure
```
[Content_Types].xml  - Defines file types in the package
_rels/.rels             - Root relationships
ppt/                     - Main presentation folder
docProps/                - Document properties
```

### PPT Folder Structure

#### 1. Main Files
- `presentation.xml` - Main presentation document
  - Slide list (sldIdLst)
  - Master slide list (sldMasterIdLst)
  - Slide size (sldSz)
  - Default text styles (defaultTextStyle)

- `_rels/presentation.xml.rels` - Relationships for presentation
  - Links to slide masters, themes, slides, etc.

#### 2. Slide Masters (slideMasters/)
- `slideMaster1.xml`, `slideMaster2.xml`, ...
- Contains:
  - Background definitions (bgRef or bgPr)
  - Common elements (footer, slide numbers)
  - Text styles (txStyles)
  - Layout list (sldLayoutIdLst)
  - Color map (clrMap)
- `_rels/slideMaster[N].xml.rels` - Master relationships

#### 3. Slide Layouts (slideLayouts/)
- `slideLayout1.xml` to `slideLayout16.xml`
- Contains:
  - Layout-specific elements
  - Placeholder definitions
  - Background definitions
- `_rels/slideLayout[N].xml.rels` - Layout relationships

#### 4. Slides (slides/)
- `slide1.xml` to `slide35.xml`
- Contains:
  - Slide elements (shapes, pictures, charts, tables)
  - Background (optional, can reference layout/master)
  - Content
- `_rels/slide[N].xml.rels` - Slide relationships to images, media, etc.

#### 5. Theme (theme/)
- `theme1.xml`
- Contains:
  - Color scheme (clrScheme)
  - Font scheme (fontScheme)
  - Effect scheme (fmtScheme)
  - Theme definitions

#### 6. Media (media/)
- `image1.emf`, `image2.jpeg`, `image3.png`, ...
- Binary image files

#### 7. Other Resources
- `embeddings/` - OLE objects (think-cell, etc.)
- `drawings/` - VML drawings
- `notesSlides/` - Speaker notes
- `notesMasters/` - Note master

## Key XML Namespaces

```xml
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
```

## Background Resolution Chain

```
1. Slide Background (slide.xml)
   ├─ <p:bgPr> - Direct background properties
   └─ <p:bgRef> - Reference to theme color

2. If not found, check SlideLayout (slideLayout.xml)
   ├─ <p:bgPr> - Direct background properties
   └─ <p:bgRef> - Reference to theme color

3. If not found, check SlideMaster (slideMaster.xml)
   ├─ <p:bgPr> - Direct background properties
   └─ <p:bgRef> - Reference to theme color

4. If not found, check Theme (theme.xml)
   └─ Theme background color
```

## Relationship References (r:embed)

Images and resources are referenced through relationships:

**In slide.xml:**
```xml
<p:pic>
  <p:blipFill>
    <a:blip r:embed="rId4"/>  <!-- Reference -->
  </p:blipFill>
</p:pic>
```

**In slide.xml.rels:**
```xml
<Relationship Id="rId4"
             Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
             Target="../media/image1.png"/>
```

## Color Scheme Colors

Theme defines these standard colors:

| Name | Description | Common Usage |
|------|-------------|--------------|
| bg1  | Background 1 | Light background |
| tx1  | Text 1       | Text on bg1 |
| bg2  | Background 2 | Dark background |
| tx2  | Text 2       | Text on bg2 |
| accent1-6 | Accent colors | Emphasis elements |
| hlink | Hyperlink | Link color |
| folHlink | Followed hyperlink | Visited link |

## Color Map Overrides

Master slides and layouts can override how theme colors are used:

```xml
<p:clrMap bg1="lt1" tx1="dk1" accent1="accent3"/>
```

This means:
- When bg1 is requested, use lt1 (light) instead
- When tx1 is requested, use dk1 (dark) instead
- When accent1 is requested, use accent3 instead

## EMU Units (English Metric Units)

PPTX uses EMU for all measurements:
- 914,400 EMU = 1 inch
- 914,400 EMU = 96 pixels (at 96 DPI)
- 360,000 EMU = 1 cm
- 1,000,000 EMU = 1 point

**Conversion:**
```
EMU to pixels:  value * 96 / 914400
Pixels to EMU:  value * 914400 / 96
```

## Sample Slide XML Structure

```xml
<p:sld>
  <p:cSld>  <!-- Common Slide Data -->
    <p:bg>  <!-- Background (optional) -->
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>  <!-- Reference theme color -->
      </p:bgRef>
    </p:bg>
    
    <p:spTree>  <!-- Shape Tree -->
      <p:sp>  <!-- Shape -->
        <p:nvSpPr>
          <p:cNvPr id="5" name="Title Placeholder"/>
          <p:nvPr>
            <p:ph type="title"/>  <!-- Placeholder type -->
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="398145" y="365125"/>
            <a:ext cx="10061184" cy="371512"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr lang="zh-CN" sz="2800"/>
              <a:t>Slide Title</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      
      <p:pic>  <!-- Picture -->
        <p:nvPicPr>
          <p:cNvPr id="14" name="Logo"/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId6"/>  <!-- Image reference -->
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="6671310"/>
            <a:ext cx="12192001" cy="291465"/>
          </a:xfrm>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
  
  <p:clrMapOvr>
    <a:masterClrMapping/>  <!-- Use master color mapping -->
  </p:clrMapOvr>
</p:sld>
```

## PPTXjs Implementation Details

### Background Parsing (Lines 11567-12133)

1. **Check slide background**
   - bgPr: Direct background properties
   - bgRef: Reference to theme

2. **Parse background fill types**
   - SOLID_FILL: Single color
   - GRADIENT_FILL: Color gradient
   - PIC_FILL: Image background

3. **Image background**
   - Extract relId from blip element
   - Get image path from relsMap
   - Read image from zip and encode to base64
   - Apply effects (duotone, tiling, etc.)

### Master/Layout Parsing (Lines 725-779)

1. **Index nodes** by id, idx, and type
   - Create lookup tables for fast access
   - Enable inheritance resolution

2. **Parse elements**
   - Extract shapes, pictures, groups
   - Resolve placeholders

3. **Color map**
   - Parse color overrides
   - Apply to theme colors

## Implementation Recommendations

### Phase 1: Core Parsing (DONE)
- ✅ Basic slide parsing
- ✅ Image extraction
- ✅ Background parsing (bgPr)
- ✅ Relationship parsing

### Phase 2: Theme & Masters (PARTIAL)
- ✅ Theme parsing
- ✅ Master slide parsing
- ✅ bgRef support
- ✅ Color scheme resolution
- ⚠️ Color map overrides
- ❌ Layout parsing
- ❌ Background inheritance chain

### Phase 3: Advanced Features
- ❌ OLE objects with fallback
- ❌ Gradient backgrounds
- ❌ Pattern fills
- ❌ VML drawings
- ❌ SmartArt full support
- ❌ Animations

### Phase 4: Rendering Enhancements
- ❌ Text formatting (multiple runs)
- ❌ Shape effects (shadows, etc.)
- ❌ Group transformations
- ❌ Connector lines
- ❌ Text fields (page numbers, dates)

## Testing Strategy

1. **Unit Tests**
   - Theme parsing
   - Master slide parsing
   - Color resolution
   - Background parsing

2. **Integration Tests**
   - Real PPTX files
   - Compare with PPTXjs output
   - Visual regression testing

3. **Performance Tests**
   - Large PPTX files
   - Many slides
   - Many images

## Reference Links

- [Office Open XML Specification](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [PPTXjs Source Code](https://github.com/meshesha/PPTXjs)
- [OpenXML Developer](https://openxmldeveloper.org/)
