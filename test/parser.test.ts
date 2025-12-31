import { describe, it, expect, beforeEach } from 'vitest';
import { parsePptx } from '../src/core';
import type { PptDocument } from '../src/types';

describe('parsePptx - PPTX解析核心功能测试', () => {
  // 创建一个简单的PPTX文件用于测试
  const createMockPptx = async (): Promise<Blob> => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();

    // 创建基本PPTX结构
    zip.file('[Content_Types].xml', `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/>
</Types>`);

    zip.file('_rels/.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

    zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:slideIdLst/>
</p:presentation>`);

    // 创建幻灯片1
    zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:grpSp>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="6858000"/>
          </a:xfrm>
        </p:spPr>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:bg>
    <p:bgPr>
      <a:solidFill>
        <a:srgbClr val="ffffff"/>
      </a:solidFill>
    </p:bgPr>
  </p:bg>
</p:sld>`);

    // 创建文档属性
    zip.file('docProps/core.xml', `<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>测试PPT</dc:title>
</cp:coreProperties>`);

    // 创建幻灯片布局
    zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?>
<sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <cSld>
    <sldSz cx="9144000" cy="6858000" type="screen"/>
  </cSld>
</sldLayout>`);

    return await zip.generateAsync({ type: 'blob' });
  };

  it('应该成功解析PPTX文件', async () => {
    const pptxBlob = await createMockPptx();
    const result = await parsePptx(pptxBlob);

    expect(result).toBeDefined();
    expect(result.id).toBeDefined();
    expect(result.title).toBe('测试PPT');
    expect(result.slides).toBeInstanceOf(Array);
  });

  it('应该正确解析文档属性', async () => {
    const pptxBlob = await createMockPptx();
    const result = await parsePptx(pptxBlob);

    expect(result.props.width).toBe(960);
    expect(result.props.height).toBe(720);
    expect(result.props.ratio).toBeCloseTo(1.33, 2);
  });

  it('应该正确解析幻灯片', async () => {
    const pptxBlob = await createMockPptx();
    const result = await parsePptx(pptxBlob);

    expect(result.slides).toHaveLength(1);
    expect(result.slides[0].id).toBeDefined();
    expect(result.slides[0].bgColor).toBe('#ffffff');
    expect(result.slides[0].elements).toBeInstanceOf(Array);
  });

  it('应该处理包含文本元素的PPTX', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();

    zip.file('[Content_Types].xml', `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`);

    zip.file('_rels/.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

    zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:slideIdLst/>
</p:presentation>`);

    zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:grpSp>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="6858000"/>
          </a:xfrm>
        </p:spPr>
        <p:sp>
          <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:rPr sz="2800" b="1"/>
                <a:t>测试文本</a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:bg>
    <p:bgPr>
      <a:solidFill>
        <a:srgbClr val="ffffff"/>
      </a:solidFill>
    </p:bgPr>
  </p:bg>
</p:sld>`);

    zip.file('docProps/core.xml', `<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>文本测试</dc:title>
</cp:coreProperties>`);

    zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?>
<sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <cSld>
    <sldSz cx="9144000" cy="6858000" type="screen"/>
  </cSld>
</sldLayout>`);

    const pptxBlob = await zip.generateAsync({ type: 'blob' });
    const result = await parsePptx(pptxBlob);

    expect(result.slides[0].elements.length).toBeGreaterThan(0);
    const textElement = result.slides[0].elements[0];
    expect(textElement.type).toBe('text');
  });

  it('应该处理多页幻灯片', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();

    zip.file('[Content_Types].xml', `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`);

    zip.file('_rels/.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

    zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:slideIdLst/>
</p:presentation>`);

    // 幻灯片1
    zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:grpSp>
        <p:spPr>
          <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="6858000"/>
          </a:xfrm>
        </p:spPr>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:bg>
    <p:bgPr>
      <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:srgbClr val="ffffff"/>
      </a:solidFill>
    </p:bgPr>
  </p:bg>
</p:sld>`);

    // 幻灯片2
    zip.file('ppt/slides/slide2.xml', `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:grpSp>
        <p:spPr>
          <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="6858000"/>
          </a:xfrm>
        </p:spPr>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:bg>
    <p:bgPr>
      <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:srgbClr val="f0f0f0"/>
      </a:solidFill>
    </p:bgPr>
  </p:bg>
</p:sld>`);

    zip.file('docProps/core.xml', `<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>多页测试</dc:title>
</cp:coreProperties>`);

    zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?>
<sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <cSld>
    <sldSz cx="9144000" cy="6858000" type="screen"/>
  </cSld>
</sldLayout>`);

    const pptxBlob = await zip.generateAsync({ type: 'blob' });
    const result = await parsePptx(pptxBlob);

    expect(result.slides).toHaveLength(2);
    expect(result.slides[0].bgColor).toBe('#ffffff');
    expect(result.slides[1].bgColor).toBe('#f0f0f0');
  });

  it('应该处理无效文件并使用默认值', async () => {
    // 创建一个最小化的无效PPTX
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();

    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types></Types>');
    zip.file('_rels/.rels', '<Relationships></Relationships>');
    zip.file('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation></p:presentation>');
    zip.file('ppt/slides/slide1.xml', '<?xml version="1.0"?><p:sld><p:cSld><p:spTree></p:spTree></p:cSld></p:sld>');

    const pptxBlob = await zip.generateAsync({ type: 'blob' });
    const result = await parsePptx(pptxBlob);

    expect(result).toBeDefined();
    expect(result.title).toBe('未命名PPT');
    expect(result.props.width).toBe(1280);
    expect(result.props.height).toBe(720);
  });
});
