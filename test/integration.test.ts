import { describe, it, expect, beforeEach } from 'vitest';
import { parsePptx, serializePptx } from '../src/core';
import type { PptDocument } from '../src/types';

describe('Integration - 集成测试', () => {
  const createCompleteMockPptx = async (): Promise<Blob> => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();

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

    // 第一页：包含文本和形状
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
                <a:t>这是第一页</a:t>
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

    // 第二页：包含表格
    zip.file('ppt/slides/slide2.xml', `<?xml version="1.0"?>
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
        <p:tbl>
          <p:tblPr>
            <a:tblW w="8000000" type="dxa"/>
          </p:tblPr>
          <a:tr>
            <a:tc>
              <a:txBody>
                <a:p>
                  <a:r>
                    <a:t>表头1</a:t>
                  </a:r>
                </a:p>
              </a:txBody>
            </a:tc>
            <a:tc>
              <a:txBody>
                <a:p>
                  <a:r>
                    <a:t>表头2</a:t>
                  </a:r>
                </a:p>
              </a:txBody>
            </a:tc>
          </a:tr>
          <a:tr>
            <a:tc>
              <a:txBody>
                <a:p>
                  <a:r>
                    <a:t>数据1</a:t>
                  </a:r>
                </a:p>
              </a:txBody>
            </a:tc>
            <a:tc>
              <a:txBody>
                <a:p>
                  <a:r>
                    <a:t>数据2</a:t>
                  </a:r>
                </a:p>
              </a:txBody>
            </a:tc>
          </a:tr>
        </p:tbl>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:bg>
    <p:bgPr>
      <a:solidFill>
        <a:srgbClr val="f0f0f0"/>
      </a:solidFill>
    </p:bgPr>
  </p:bg>
</p:sld>`);

    zip.file('docProps/core.xml', `<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>集成测试PPT</dc:title>
</cp:coreProperties>`);

    zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?>
<sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <cSld>
    <sldSz cx="9144000" cy="6858000" type="screen"/>
  </cSld>
</sldLayout>`);

    return await zip.generateAsync({ type: 'blob' });
  };

  it('应该完整解析多页PPTX', async () => {
    const pptxBlob = await createCompleteMockPptx();
    const result = await parsePptx(pptxBlob);

    expect(result.title).toBe('集成测试PPT');
    expect(result.slides.length).toBeGreaterThanOrEqual(2);
    expect(result.props.width).toBe(960);
    expect(result.props.height).toBe(720);
  });

  it('应该解析并序列化保持基本结构', async () => {
    const pptxBlob = await createCompleteMockPptx();
    const parsedDoc = await parsePptx(pptxBlob);

    const serializedBlob = await serializePptx(parsedDoc);
    expect(serializedBlob).toBeInstanceOf(Blob);

    // 重新解析序列化的文件
    const reParsedDoc = await parsePptx(serializedBlob);
    expect(reParsedDoc.title).toBe(parsedDoc.title);
    expect(reParsedDoc.slides.length).toBe(parsedDoc.slides.length);
  });

  it('应该正确处理中文内容', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();

    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
    zip.file('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:slideIdLst/></p:presentation>');

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

    zip.file('docProps/core.xml', `<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>中文测试文档</dc:title>
</cp:coreProperties>`);

    zip.file('ppt/slideLayouts/slideLayout1.xml', '<?xml version="1.0"?><sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main"><cSld><sldSz cx="9144000" cy="6858000" type="screen"/></cSld></sldLayout>');

    const pptxBlob = await zip.generateAsync({ type: 'blob' });
    const result = await parsePptx(pptxBlob);

    expect(result.title).toBe('中文测试文档');
  });

  it('应该处理大文档', async () => {
    const JSZip = (await import('jszip')).default;

    // 创建一个包含多张幻灯片的文档
    const createLargeDoc = async (): Promise<Blob> => {
      const zip = new JSZip();

      zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
      zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
      zip.file('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:slideIdLst/></p:presentation>');

      for (let i = 1; i <= 5; i++) {
        zip.file(`ppt/slides/slide${i}.xml`, `<?xml version="1.0"?>
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
      }

      zip.file('docProps/core.xml', '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>大文档测试</dc:title></cp:coreProperties>');
      zip.file('ppt/slideLayouts/slideLayout1.xml', '<?xml version="1.0"?><sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main"><cSld><sldSz cx="9144000" cy="6858000" type="screen"/></cSld></sldLayout>');

      return await zip.generateAsync({ type: 'blob' });
    };

    const pptxBlob = await createLargeDoc();
    const result = await parsePptx(pptxBlob);

    expect(result.slides.length).toBe(5);

    // 序列化大文档
    const serializedBlob = await serializePptx(result);
    expect(serializedBlob).toBeInstanceOf(Blob);

    const reParsedDoc = await parsePptx(serializedBlob);
    expect(reParsedDoc.slides.length).toBe(5);
  });

  it('应该正确处理不同比例的幻灯片', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();

    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
    zip.file('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:slideIdLst/></p:presentation>');

    zip.file('ppt/slides/slide1.xml', '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:grpSp><p:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/></a:xfrm></p:spPr></p:grpSp></p:spTree></p:cSld></p:sld>');

    zip.file('docProps/core.xml', '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>16:9比例</dc:title></cp:coreProperties>');

    // 16:9 比例
    zip.file('ppt/slideLayouts/slideLayout1.xml', '<?xml version="1.0"?><sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main"><cSld><sldSz cx="9144000" cy="5143500" type="screen"/></cSld></sldLayout>');

    const pptxBlob = await zip.generateAsync({ type: 'blob' });
    const result = await parsePptx(pptxBlob);

    expect(result.props.width).toBe(960);
    expect(result.props.height).toBe(540);
    expect(result.props.ratio).toBeCloseTo(1.78, 2);
  });
});
