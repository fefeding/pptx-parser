/**
 * Mock PPTX生成器 - 扩展版
 * 支持富文本、布局关联等复杂结构
 * 
 * 对齐PPTXjs的XML结构要求
 */

export class MockPptxGenerator {
  private zip: any;
  private slideCount = 0;
  private layoutCount = 0;
  private mediaCount = 0;

  constructor() {
    // 动态导入JSZip
    this.initZip();
  }

  private async initZip() {
    const JSZip = (await import('jszip')).default;
    this.zip = new JSZip();
  }

  /**
   * 创建基础PPTX结构
   */
  async createBaseStructure(): Promise<void> {
    await this.initZip();

    // Content Types
    this.zip.file('[Content_Types].xml', `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`);

    // Root relationships
    this.zip.file('_rels/.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

    // Presentation.xml
    this.zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:slideIdLst/>
</p:presentation>`);

    // 主题文件
    this.zip.file('ppt/theme/theme1.xml', `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Arial"/>
        <a:ea typeface="微软雅黑"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Arial"/>
        <a:ea typeface="微软雅黑"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`);

    // 文档属性
    this.zip.file('docProps/core.xml', `<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>测试PPT</dc:title>
</cp:coreProperties>`);
  }

  /**
   * 创建布局文件
   * 对齐PPTXjs的slideLayout结构
   */
  async createLayout(layoutId: number, options: {
    titlePlaceholder?: { x: number; y: number; width: number; height: number };
    bodyPlaceholder?: { x: number; y: number; width: number; height: number };
  } = {}): Promise<string> {
    this.layoutCount++;
    const layoutFileName = `ppt/slideLayouts/slideLayout${layoutId}.xml`;
    
    let placeholdersXml = '';
    
    // 标题占位符
    if (options.titlePlaceholder) {
      placeholdersXml += `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="1" name="Title 1"/>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${options.titlePlaceholder.x * 9144}" y="${options.titlePlaceholder.y * 9144}"/>
            <a:ext cx="${options.titlePlaceholder.width * 9144}" cy="${options.titlePlaceholder.height * 9144}"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
          </a:p>
        </p:txBody>
      </p:sp>`;
    }
    
    // 正文占位符
    if (options.bodyPlaceholder) {
      placeholdersXml += `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Content Placeholder 1"/>
          <p:nvPr>
            <p:ph idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${options.bodyPlaceholder.x * 9144}" y="${options.bodyPlaceholder.y * 9144}"/>
            <a:ext cx="${options.bodyPlaceholder.width * 9144}" cy="${options.bodyPlaceholder.height * 9144}"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="l"/>
          </a:p>
        </p:txBody>
      </p:sp>`;
    }

    this.zip.file(layoutFileName, `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>${placeholdersXml}
    </p:spTree>
  </p:cSld>
</p:sldLayout>`);

    // 布局关系文件
    const layoutRelsFile = `ppt/slideLayouts/_rels/slideLayout${layoutId}.xml.rels`;
    this.zip.file(layoutRelsFile, `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);

    return layoutFileName;
  }

  /**
   * 创建带富文本的幻灯片
   * 对齐PPTXjs的文本结构要求
   */
  async createRichTextSlide(slideData: {
    title?: string;
    content?: Array<{
      text: string;
      fontSize?: number;
      color?: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
    }>;
    backgroundColor?: string;
    layoutId?: number;
  }): Promise<string> {
    this.slideCount++;
    const slideId = this.slideCount;
    const slideFileName = `ppt/slides/slide${slideId}.xml`;

    let bodyContent = '';
    
    if (slideData.content && slideData.content.length > 0) {
      slideData.content.forEach((run, index) => {
        const fontSize = run.fontSize || 18; // 默认18pt
        const fontSizeEmu = fontSize * 100; // 转换为PPTX单位
        
        let styleAttrs = '';
        if (run.bold) styleAttrs += ' b="1"';
        if (run.italic) styleAttrs += ' i="1"';
        if (run.underline) styleAttrs += ' u="sng"';
        
        bodyContent += `
            <a:p>
              <a:r>
                <a:rPr sz="${fontSizeEmu}"${styleAttrs}>
                  <a:solidFill>
                    <a:srgbClr val="${run.color || '000000'}"/>
                  </a:solidFill>
                  <a:latin typeface="Arial"/>
                </a:rPr>
                <a:t>${this.escapeXml(run.text)}</a:t>
              </a:r>
            </a:p>`;
      });
    }

    let titleContent = '';
    if (slideData.title) {
      const titleFontSize = 44; // 标题默认44pt
      const fontSizeEmu = titleFontSize * 100;
      
      titleContent = `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Title 1"/>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="4572000" y="2743200"/>
            <a:ext cx="9144000" cy="1371600"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr sz="${fontSizeEmu}" b="1">
                <a:solidFill>
                  <a:srgbClr val="000000"/>
                </a:solidFill>
                <a:latin typeface="Arial"/>
              </a:rPr>
              <a:t>${this.escapeXml(slideData.title)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`;
    }

    this.zip.file(slideFileName, `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Text Box 1"/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="7315200" cy="5029200"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${bodyContent}
        </p:txBody>
      </p:sp>${titleContent}
    </p:spTree>
  </p:cSld>
  <p:bg>
    <p:bgPr>
      <a:solidFill>
        <a:srgbClr val="${slideData.backgroundColor || 'ffffff'}"/>
      </a:solidFill>
    </p:bgPr>
  </p:bg>
</p:sld>`);

    // 幻灯片关系文件
    const slideRelsFile = `ppt/slides/_rels/slide${slideId}.xml.rels`;
    const layoutId = slideData.layoutId || 1;
    this.zip.file(slideRelsFile, `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${layoutId}.xml"/>
</Relationships>`);

    return slideFileName;
  }

  /**
   * 添加图片媒体资源
   * 对齐PPTXjs的图片处理逻辑
   */
  async addImage(imageData: {
    fileName: string;
    mimeType: string;
    data: string; // base64编码
  }): Promise<string> {
    this.mediaCount++;
    const mediaFileName = `ppt/media/image${this.mediaCount}.${imageData.fileName.split('.').pop()}`;
    
    // 将base64转换为二进制数据
    const binaryString = atob(imageData.data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    
    this.zip.file(mediaFileName, bytes);
    return mediaFileName;
  }

  /**
   * 创建带图片的幻灯片
   * 对齐PPTXjs的图片结构要求
   */
  async createImageSlide(slideData: {
    images: Array<{
      mediaPath: string;
      x: number;
      y: number;
      width: number;
      height: number;
    }>;
    backgroundColor?: string;
  }): Promise<string> {
    this.slideCount++;
    const slideId = this.slideCount;
    const slideFileName = `ppt/slides/slide${slideId}.xml`;

    let imagesXml = '';
    slideData.images.forEach((image, index) => {
      // 转换位置为EMU单位
      const xEmu = image.x * 9144;
      const yEmu = image.y * 9144;
      const widthEmu = image.width * 9144;
      const heightEmu = image.height * 9144;

      imagesXml += `
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="${4 + index}" name="Picture ${index + 1}"/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId${index + 1}"/>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="${xEmu}" y="${yEmu}"/>
            <a:ext cx="${widthEmu}" cy="${heightEmu}"/>
          </a:xfrm>
        </p:spPr>
      </p:pic>`;
    });

    this.zip.file(slideFileName, `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>${imagesXml}
    </p:spTree>
  </p:cSld>
  <p:bg>
    <p:bgPr>
      <a:solidFill>
        <a:srgbClr val="${slideData.backgroundColor || 'ffffff'}"/>
      </a:solidFill>
    </p:bgPr>
  </p:bg>
</p:sld>`);

    // 幻灯片关系文件
    const slideRelsFile = `ppt/slides/_rels/slide${slideId}.xml.rels`;
    let slideRelsXml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    slideData.images.forEach((image, index) => {
      slideRelsXml += `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${image.mediaPath.split('/').pop()}"/>`;
    });
    slideRelsXml += '</Relationships>';

    this.zip.file(slideRelsFile, slideRelsXml);

    return slideFileName;
  }

  /**
   * 生成PPTX文件
   */
  async generate(): Promise<Blob> {
    // 更新presentation.xml中的幻灯片列表
    let slideIdListXml = '<p:sldIdLst>';
    for (let i = 1; i <= this.slideCount; i++) {
      slideIdListXml += `<p:sldId id="${256 + i}" rId="rId${i}"/>`;
    }
    slideIdListXml += '</p:sldIdLst>';

    this.zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:sldSz cx="9144000" cy="6858000"/>
  ${slideIdListXml}
</p:presentation>`);

    // 更新presentation.xml的rels
    let presentationRelsXml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    presentationRelsXml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="slideLayouts/slideLayout1.xml"/>';
    presentationRelsXml += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
    for (let i = 1; i <= this.slideCount; i++) {
      presentationRelsXml += `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i}.xml"/>`;
    }
    presentationRelsXml += '</Relationships>';

    this.zip.file('ppt/_rels/presentation.xml.rels', presentationRelsXml);

    return await this.zip.generateAsync({ type: 'blob' });
  }

  private escapeXml(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}

/**
 * 创建富文本PPTX的便捷函数
 */
export async function createRichTextPptx(options: {
  title?: string;
  content?: Array<{
    text: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
  }>;
  backgroundColor?: string;
}): Promise<Blob> {
  const generator = new MockPptxGenerator();
  await generator.createBaseStructure();
  await generator.createLayout(1);
  await generator.createRichTextSlide(options);
  return await generator.generate();
}

/**
 * 创建带图片PPTX的便捷函数
 */
export async function createImagePptx(options: {
  images: Array<{
    fileName: string;
    mimeType: string;
    data: string; // base64
    x: number;
    y: number;
    width: number;
    height: number;
  }>;
  backgroundColor?: string;
}): Promise<Blob> {
  const generator = new MockPptxGenerator();
  await generator.createBaseStructure();

  // 添加所有图片
  const mediaPaths = await Promise.all(options.images.map(async (img) => {
    return await generator.addImage({
      fileName: img.fileName,
      mimeType: img.mimeType,
      data: img.data
    });
  }));

  // 创建幻灯片
  await generator.createImageSlide({
    images: options.images.map((img, index) => ({
      mediaPath: mediaPaths[index],
      x: img.x,
      y: img.y,
      width: img.width,
      height: img.height
    })),
    backgroundColor: options.backgroundColor
  });

  return await generator.generate();
}