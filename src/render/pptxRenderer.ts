import { PptxToHtmlOptions } from "../types/index";
import { PptxParser } from "../core/pptxParser";

/** PPTX转HTML渲染器 - 纯字符串返回，保持Node与浏览器接口一致 */
export class PptxRenderer {
  private parser: PptxParser;
  private styleTable: Record<string, { name: string; text: string }> = {};

  constructor(private options: PptxToHtmlOptions, private buffer: Uint8Array) {
    this.parser = new PptxParser(buffer);
  }

  /** 核心渲染方法 - 始终返回HTML字符串 */
  async render(): Promise<string> {
    //1. 初始化解析器
    await this.parser.init();
    const slides = this.parser.getSlides();
    const totalSlides = this.parser.getTotalSlides();
    const slideWidth = this.parser.getSlideWidth();
    const slideHeight = this.parser.getSlideHeight();

    // 2. 渲染幻灯片容器
    let html = '<div class="pptxjs-container">';

    // 3. 逐页渲染幻灯片
    for (let i = 1; i <= totalSlides; i++) {
      const slideData = slides[i];
      if (slideData) {
        // 构建warpObj，包含layout、master和theme信息
        const warpObj = this.parser.getWarpObjForSlide(i);
        html += this.renderSlide(slideData, i, slideWidth, slideHeight, warpObj);
      }
    }

    // 添加全局CSS
    const globalCSS = this.generateGlobalCSS();
    if (globalCSS) {
      html += `<style>${globalCSS}</style>`;
    }

    html += '</div>';
    return html;
  }

  /** 渲染单张幻灯片 - 返回HTML字符串 */
  private renderSlide(slideData: any, slideId: number, slideWidth: number, slideHeight: number, warpObj: any): string {
    const isRevealJs = this.options.slideType === "revealjs";
    const tag = isRevealJs ? "section" : "div";

    let html = `<${tag} class="slide" data-slide-id="${slideId}" style="width:${slideData.width || slideWidth}px;height:${slideData.height || slideHeight}px;">`;

    // 渲染背景（颜色或图片）- 使用slideData中已经处理好的背景
    if (slideData.backgroundColor && !slideData.backgroundImage && !slideData.backgroundImageData) {
      html += `<div class="slide-background" style="width:100%;height:100%;background-color:${slideData.backgroundColor};"></div>`;
    } else if (slideData.backgroundImageData) {
      html += `<div class="slide-background" style="width:100%;height:100%;background-image:url('${slideData.backgroundImageData}');background-size:cover;background-position:center;"></div>`;
    }

    // 渲染layout中的占位符内容（通过slideData.layout.shapes）
    if (slideData.layout && slideData.layout.shapes) {
      slideData.layout.shapes.forEach((shape: any) => {
        // 只渲染layout中未在slide中覆盖的占位符
        const isOverridden = slideData.shapes && slideData.shapes.some((slideShape: any) => 
          slideShape.placeholder?.idx === shape.placeholderType?.placeholder?.idx
        );
        if (!isOverridden) {
          html += this.renderShape(shape, warpObj);
        }
      });
    }

    // 渲染master中的占位符内容（通过slideData.master.shapes）
    if (slideData.master && slideData.master.shapes) {
      slideData.master.shapes.forEach((shape: any) => {
        // 只渲染master中未在layout和slide中覆盖的占位符
        const isLayoutOverridden = slideData.layout && slideData.layout.shapes && 
          slideData.layout.shapes.some((layoutShape: any) => 
            layoutShape.placeholder?.idx === shape.placeholderType?.placeholder?.idx
          );
        const isSlideOverridden = slideData.shapes && slideData.shapes.some((slideShape: any) => 
          slideShape.placeholder?.idx === shape.placeholderType?.placeholder?.idx
        );
        if (!isLayoutOverridden && !isSlideOverridden) {
          html += this.renderShape(shape, warpObj);
        }
      });
    }

    // 渲染形状（包括带文本的形状）
    if (slideData.shapes) {
      slideData.shapes.forEach((shape: any) => {
        html += this.renderShape(shape, warpObj);
      });
    }

    // 渲染图片
    if (slideData.images) {
      slideData.images.forEach((image: any) => {
        html += this.renderImage(image);
      });
    }

    // 渲染表格
    if (slideData.tables) {
      slideData.tables.forEach((table: any) => {
        html += this.renderTable(table);
      });
    }

    // 渲染图表
    if (slideData.graphs) {
      slideData.graphs.forEach((graph: any) => {
        html += this.renderChart(graph);
      });
    }

    html += `</${tag}>`;
    return html;
  }

  /** 渲染形状 */
  private renderShape(shape: any, warpObj?: any): string {
    const style: string[] = [
      `top:${shape.y || 0}px`,
      `left:${shape.x || 0}px`,
      `width:${shape.width || 0}px`,
      `height:${shape.height || 0}px`,
      `1px solid hidden`,
      `background-color: inherit`,
      `z-index: ${shape.id || 1}`,
      `transform: rotate(${shape.rotation || 0}deg)`,
    ];

    let html = `<div class="block v-up content"`;

    // 添加属性
    if (shape.id !== undefined) html += ` _id="${shape.id}"`;
    if (shape.placeholder?.idx !== undefined) html += ` _idx="${shape.placeholder.idx}"`;
    if (shape.placeholder?.type) html += ` _type="${shape.placeholder.type}"`;
    if (shape.name) html += ` _name="${shape.name}"`;

    html += ` style="${style.join(';')}">`;

    // 获取文本体节点 - 优先从shape获取,其次从node获取
    let textBodyNode: any = undefined;
    if (shape.textBodyNode) {
      textBodyNode = shape.textBodyNode;
    } else if (shape.node) {
      textBodyNode = shape.node["p:txBody"];
    }

    // 渲染文本内容 - 传入完整的warpObj以支持样式继承
    if (textBodyNode && warpObj) {
      html += this.genTextBody(
        textBodyNode,
        shape,
        warpObj,
        shape.type,
        shape.placeholder?.idx
      );
    } else if (shape.textBlocks) {
      // 回退到旧逻辑
      html += this.renderTextBlocksLegacy(shape.textBlocks, shape);
    }

    html += '</div>';
    return html;
  }

  /** 生成文本主体HTML - 模仿PPTXjs的genTextBody */
  private genTextBody(
    textBodyNode: any,
    spNode: any,
    warpObj: any,
    type: string,
    idx?: number
  ): string {
    let text = "";

    if (!textBodyNode) {
      return text;
    }

    const pFontStyle = this.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
    const apNode = textBodyNode["a:p"];
    const paragraphs = Array.isArray(apNode) ? apNode : [apNode];

    for (let i = 0; i < paragraphs.length; i++) {
      const pNode = paragraphs[i];
      
      // 生成段落样式
      let styleText = "";
      
      // 获取垂直边距
      const marginsVer = this.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
      if (marginsVer) {
        styleText = marginsVer;
      }

      // 特定类型的默认样式
      if (type === "body" || type === "obj" || type === "shape") {
        styleText += "font-size: 0px;";
        styleText += "font-weight: 100;";
        styleText += "font-style: normal;";
      }

      // 生成段落容器的CSS类名
      const cssName = this.getStyleClass(styleText);

      // 获取段落宽度
      const prg_width_node = this.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
      const prg_width = prg_width_node !== undefined ? `width:${parseInt(prg_width_node) * (96 / 914400)}px;` : "width:inherit;";

      // 获取段落方向
      const prg_dir = this.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);

      text += `<div style='display: flex;${prg_width}' class='slide-prgrph h-left pregraph-ltr ${cssName} ${prg_dir}'>`;
      
      // 生成文本内容容器
      text += `<div style='height:100%;direction:initial;overflow-wrap:break-word;word-wrap:break-word;${prg_width}'>`;

      // 处理段落中的文本运行（a:r）
      let rNode = pNode["a:r"];
      let fldNode = pNode["a:fld"];
      let brNode = pNode["a:br"];

      if (rNode !== undefined) {
        rNode = Array.isArray(rNode) ? rNode : [rNode];
      }
      if (fldNode !== undefined) {
        fldNode = Array.isArray(fldNode) ? fldNode : [fldNode];
      }

      if (rNode !== undefined && fldNode !== undefined) {
        rNode = rNode.concat(fldNode);
      }

      if (rNode !== undefined) {
        // 为每个文本运行生成span
        for (let j = 0; j < rNode.length; j++) {
          text += this.genSpanElement(
            rNode[j],
            j,
            pNode,
            textBodyNode,
            pFontStyle,
            undefined,
            idx,
            type,
            rNode.length,
            warpObj
          );
        }
      }

      text += "</div></div>";
    }

    return text;
  }

  /** 渲染文本块 - 旧版逻辑（回退） */
  private renderTextBlocksLegacy(textBlocks: any[], shape?: any): string {
    let html = '';

    if (textBlocks && textBlocks.length > 0) {
      textBlocks.forEach((block: any) => {
        if (block.runs && block.runs.length > 0) {
          block.runs.forEach((run: any) => {
            const styleParts: string[] = [];
            
            if (run.fontSize !== undefined) {
              styleParts.push(`font-size:${run.fontSize}px`);
            }
            
            if (run.fontFamily) {
              styleParts.push(`font-family:${run.fontFamily}`);
            }
            
            styleParts.push(`font-weight:${run.bold ? 'bold' : 'inherit'}`);
            styleParts.push(`font-style:${run.italic ? 'italic' : 'inherit'}`);
            styleParts.push(`text-decoration:${run.underline ? 'underline' : run.strike ? 'line-through' : 'inherit'}`);
            styleParts.push(`text-align:left`);
            styleParts.push(`vertical-align:baseline`);
            
            if (run.color) {
              styleParts.push(`color:${run.color}`);
            }

            const styleText = styleParts.join(';');
            const className = this.getStyleClass(styleText);
            
            html += `<span class="text-block ${className}">${this.escapeHtml(run.text || '')}</span>`;
          });
        }
      });
    }

    return html;
  }

  /** 生成span元素 - 模仿PPTXjs的genSpanElement */
  private genSpanElement(
    node: any,
    rIndex: number,
    pNode: any,
    textBodyNode: any,
    pFontStyle: any,
    slideLayoutSpNode: any,
    idx: number | undefined,
    type: string,
    rNodeLength: number,
    warpObj: any,
    isBullate = false
  ): string {
    let text_style = "";
    const lstStyle = textBodyNode ? textBodyNode["a:lstStyle"] : undefined;
    const slideMasterTextStyles = warpObj["slideMasterTextStyles"];

    let text = node["a:t"];
    if (typeof text !== 'string') {
      text = this.getTextByPathList(node, ["a:fld", "a:t"]);
      if (typeof text !== 'string') {
        text = "&nbsp;";
      }
    }

    // 获取级别
    const pPrNode = pNode["a:pPr"];
    let lvl = 1;
    const lvlNode = this.getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
      lvl = parseInt(lvlNode) + 1;
    }

    // 获取字体大小
    const font_size = this.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
    text_style += `font-size:${font_size};`;
    
    // 获取字体族
    const font_family = this.getFontType(node, type, warpObj, pFontStyle);
    text_style += `font-family:${font_family};`;
    
    // 获取字体粗细
    const font_weight = this.getFontBold(node, type, slideMasterTextStyles);
    text_style += `font-weight:${font_weight};`;
    
    // 获取字体样式
    const font_style = this.getFontItalic(node, type, slideMasterTextStyles);
    text_style += `font-style:${font_style};`;
    
    // 获取文本装饰
    const text_decoration = this.getFontDecoration(node, type, slideMasterTextStyles);
    text_style += `text-decoration:${text_decoration};`;
    
    // 获取文本对齐
    const text_align = this.getTextHorizontalAlign(node, pNode, type, warpObj);
    text_style += `text-align:${text_align};`;
    
    // 获取垂直对齐
    const vertical_align = this.getTextVerticalAlign(node, type, slideMasterTextStyles);
    text_style += `vertical-align:${vertical_align};`;

    // 获取字体颜色
    const fontColorPr = this.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
    const fontClrType = fontColorPr[2];
    if (fontClrType === "solid") {
      if (fontColorPr[0] !== undefined && fontColorPr[0] !== "") {
        text_style += `color:#${fontColorPr[0]};`;
      }
      if (fontColorPr[1] !== undefined && fontColorPr[1] !== "") {
        text_style += `text-shadow:${fontColorPr[1]};`;
      }
      if (fontColorPr[3] !== undefined && fontColorPr[3] !== "") {
        text_style += `background-color:#${fontColorPr[3]};`;
      }
    }

    // 样式文本（用于CSS类）- 包含垂直边距等复杂样式，模仿PPTXjs的完整样式
    let styleText = this.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
    
    // 添加字体相关样式到CSS类（PPTXjs将所有样式都放在CSS类中）
    if (font_size !== "inherit" && font_size !== "0px") {
      styleText += `font-size:${font_size};`;
    }
    
    if (font_family !== "inherit") {
      styleText += `font-family:${font_family};`;
    }
    
    if (font_weight !== "inherit" && font_weight !== "normal") {
      styleText += `font-weight:${font_weight};`;
    }
    
    if (font_style !== "inherit" && font_style !== "normal") {
      styleText += `font-style:${font_style};`;
    }
    
    if (text_decoration !== "inherit" && text_decoration !== "none") {
      styleText += `text-decoration:${text_decoration};`;
    }
    
    if (text_align !== "left") {
      styleText += `text-align:${text_align};`;
    }
    
    if (vertical_align !== "baseline") {
      styleText += `vertical-align:${vertical_align};`;
    }

    // RTL处理
    const lang = this.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
    const rtl_langs_array = ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR", "ur-PK"];
    const isRtlLan = (lang !== undefined && rtl_langs_array.indexOf(lang) !== -1);
    
    if (isRtlLan) {
      styleText += "direction:rtl;";
    } else {
      styleText += "direction:ltr;";
    }

    // 添加颜色样式到CSS类
    if (fontClrType === "solid") {
      if (fontColorPr[0] !== undefined && fontColorPr[0] !== "") {
        styleText += `color:#${fontColorPr[0]};`;
      }
    }

    // 生成CSS类名（不包含text-block前缀）
    const cssName = this.getStyleClass(styleText);

    // 内联样式只保留颜色等特殊效果，字体大小等都在CSS类中
    let inlineStyle = "";
    if (fontClrType === "solid") {
      if (fontColorPr[1] !== undefined && fontColorPr[1] !== "") {
        inlineStyle += `text-shadow:${fontColorPr[1]};`;
      }
      if (fontColorPr[3] !== undefined && fontColorPr[3] !== "") {
        inlineStyle += `background-color:#${fontColorPr[3]};`;
      }
    }

    return `<span class="${cssName}" style="${inlineStyle}">${this.escapeHtml(text || '')}</span>`;
  }

  /** 获取字体大小 - 模仿PPTXjs的getFontSize */
  private getFontSize(node: any, textBodyNode: any, pFontStyle: any, lvl: number, type: string, warpObj: any): string {
    const fontSizeFactor = 4 / 3.2;
    let fontSize: number | undefined = undefined;
    let sz: any, kern: any;
    const lstStyle = textBodyNode ? textBodyNode["a:lstStyle"] : undefined;
    const lvlpPr = `a:lvl${lvl}pPr`;

    if (node["a:rPr"] !== undefined) {
      const szAttr = this.getAttr(node["a:rPr"], "sz");
      if (szAttr !== undefined) {
        fontSize = parseInt(szAttr) / 100;
      }
    }

    if (typeof fontSize !== "number" && node["a:fld"] !== undefined) {
      sz = this.getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
      if (sz !== undefined) {
        fontSize = parseInt(sz) / 100;
      }
    }

    if (typeof fontSize !== "number" && node["a:t"] === undefined) {
      // endParaRPr不存在于span级别，跳过
    }

    if (typeof fontSize !== "number" && lstStyle !== undefined) {
      sz = this.getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
      if (sz !== undefined) {
        fontSize = parseInt(sz) / 100;
      }
    }

    // 从layout获取
    if (typeof fontSize !== "number") {
      sz = this.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
      kern = this.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
      if (sz !== undefined) {
        fontSize = parseInt(sz) / 100;
      }
      if (kern !== undefined && typeof fontSize === "number" && (fontSize - parseInt(kern) / 100) > 0) {
        fontSize = fontSize - parseInt(kern) / 100;
      }
    }

    // 从master获取
    if (typeof fontSize !== "number") {
      sz = this.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
      kern = this.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
      if (sz === undefined) {
        if (type === "title" || type === "subTitle" || type === "ctrTitle") {
          sz = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
          kern = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
        } else if (type === "body" || type === "obj" || type === "dt" || type === "sldNum" || type === "textBox") {
          sz = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
          kern = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
        } else if (type === "shape") {
          sz = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
          kern = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
        }
      }
      if (sz !== undefined) {
        fontSize = parseInt(sz) / 100;
      }
      if (kern !== undefined && typeof fontSize === "number" && ((fontSize - parseInt(kern) / 100) > parseInt(kern) / 100)) {
        fontSize = fontSize - parseInt(kern) / 100;
      }
    }

    if (typeof fontSize !== "number") {
      sz = this.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "sz"]);
      if (sz !== undefined) {
        fontSize = parseInt(sz) / 100;
      }
    }

    // 设置默认字体大小，避免0px
    if (typeof fontSize !== "number" || fontSize <= 0) {
      // 根据类型设置合理的默认值
      if (type === "title" || type === "subTitle" || type === "ctrTitle") {
        fontSize = 44; // 标题默认44pt
      } else if (type === "body" || type === "obj" || type === "dt" || type === "sldNum") {
        fontSize = 18; // 正文默认18pt
      } else {
        fontSize = 14; // 其他默认14pt
      }
    }

    return `${fontSize * fontSizeFactor}px`;
  }

  /** 获取字体类型 - 模仿PPTXjs的getFontType */
  private getFontType(node: any, type: string, warpObj: any, pFontStyle: any): string {
    // 优先级1: 检查 a:ea (东亚字体)
    let typeface = this.getTextByPathList(node, ["a:rPr", "a:ea", "attrs", "typeface"]);

    // 优先级2: 检查 a:latin (拉丁字体)
    if (typeface === undefined) {
      typeface = this.getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);
    }

    // 优先级3: 检查 a:cs (复杂脚本字体)
    if (typeface === undefined) {
      typeface = this.getTextByPathList(node, ["a:rPr", "a:cs", "attrs", "typeface"]);
    }

    // 如果没有找到字体,从主题中获取
    if (typeface === undefined && warpObj && warpObj["themeContent"]) {
      let fontIdx: string | undefined = undefined;
      let fontGrup = "";
      if (pFontStyle && pFontStyle.attrs) {
        fontIdx = this.getAttr(pFontStyle, "idx");
      }

      const fontSchemeNode = this.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
      if (fontSchemeNode) {
        if (!fontIdx) {
          if (type === "title" || type === "subTitle" || type === "ctrTitle") {
            fontIdx = "major";
          } else {
            fontIdx = "minor";
          }
        }
        fontGrup = `a:${fontIdx}Font`;

        // 优先级1: 东亚字体
        typeface = this.getTextByPathList(fontSchemeNode, [fontGrup, "a:ea", "attrs", "typeface"]);

        // 优先级2: 拉丁字体
        if (typeface === undefined) {
          typeface = this.getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
        }

        // 优先级3: 复杂脚本字体
        if (typeface === undefined) {
          typeface = this.getTextByPathList(fontSchemeNode, [fontGrup, "a:cs", "attrs", "typeface"]);
        }
      }
    }

    return (typeface === undefined || typeface === "") ? "inherit" : typeface;
  }

  /** 获取字体粗细 - 模仿PPTXjs的getFontBold */
  private getFontBold(node: any, type: string, slideMasterTextStyles: any): string {
    return (node["a:rPr"] !== undefined && this.getAttr(node["a:rPr"], "b") === "1") ? "bold" : "inherit";
  }

  /** 获取字体样式 - 模仿PPTXjs的getFontItalic */
  private getFontItalic(node: any, type: string, slideMasterTextStyles: any): string {
    return (node["a:rPr"] !== undefined && this.getAttr(node["a:rPr"], "i") === "1") ? "italic" : "inherit";
  }

  /** 获取文本装饰 - 模仿PPTXjs的getFontDecoration */
  private getFontDecoration(node: any, type: string, slideMasterTextStyles: any): string {
    if (node["a:rPr"] !== undefined) {
      const underLine = this.getAttr(node["a:rPr"], "u");
      const strikethrough = this.getAttr(node["a:rPr"], "strike");
      
      if (underLine !== undefined && underLine !== "none" && strikethrough === undefined) {
        return "underline";
      }
      if (underLine === undefined && strikethrough !== undefined) {
        return "line-through";
      }
      if (underLine !== undefined && strikethrough !== undefined) {
        return "underline line-through";
      }
    }
    return "none";
  }

  /** 获取文本水平对齐 - 模仿PPTXjs的getTextHorizontalAlign */
  private getTextHorizontalAlign(node: any, pNode: any, type: string, warpObj: any): string {
    let alignNode = this.getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);

    const lvlNode = this.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
    const lvlIdx = lvlNode !== undefined ? parseInt(lvlNode) + 1 : 1;
    const lvlStr = `a:lvl${lvlIdx}pPr`;

    const lstStyle = pNode ? this.getTextByPathList(pNode["p:txBody"], ["a:lstStyle"]) : undefined;

    // 从 layout 和 master 继承
    if (alignNode === undefined) {
      alignNode = this.getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);
    }

    if (alignNode === undefined && warpObj["slideLayoutTables"]) {
      const idxTable = warpObj["slideLayoutTables"]["idxTable"];
      if (idxTable) {
        // 注意:这里简化了,实际应该根据shape的idx来查找
        const layoutKey = Object.keys(idxTable)[0];
        if (layoutKey) {
          alignNode = this.getTextByPathList(idxTable[layoutKey], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
        }
      }
    }

    if (alignNode === undefined && warpObj["slideMasterTextStyles"]) {
      if (type === "title" || type === "ctrTitle") {
        alignNode = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlStr, "attrs", "algn"]);
      } else if (type === "body" || type === "obj" || type === "subTitle") {
        alignNode = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlStr, "attrs", "algn"]);
      } else if (type === "shape") {
        alignNode = this.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlStr, "attrs", "algn"]);
      }
    }

    switch (alignNode) {
      case "ctr": return "center";
      case "r": return "right";
      case "just": return "justify";
      case "justLow": return "justify";
      case "dist": return "justify";
      case "thaiDist": return "justify";
      default: return "left";
    }
  }

  /** 获取文本对齐 - 模仿PPTXjs的getTextVerticalAlign */
  private getTextVerticalAlign(node: any, type: string, slideMasterTextStyles: any): string {
    let anchor = this.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
    
    // 简化版 - 不回退到layout/master
    return anchor || "baseline";
  }

  /** 获取段落方向 */
  private getPregraphDir(pNode: any, textBodyNode: any, idx: number | undefined, type: string, warpObj: any): string {
    // 简化版
    return "pregraph-ltr";
  }

  /** 获取垂直边距 */
  private getVerticalMargins(pNode: any, textBodyNode: any, type: string, idx: number | undefined, warpObj: any): string {
    // margin-top: a:pPr => a:spcBef => a:spcPts (/100) | a:spcPct (/?)
    // margin-bottom: a:pPr => a:spcAft => a:spcPts (/100) | a:spcPct (/?)
    // line spacing: a:pPr => a:lnSpc => a:spcPts (/?) | a:spcPct (/?)

    let lvl = 1;
    let spcBefNode = this.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
    let spcAftNode = this.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
    let lnSpcNode = this.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
    let lnSpcNodeType = "Pct";

    if (lnSpcNode === undefined) {
      lnSpcNode = this.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
      if (lnSpcNode !== undefined) {
        lnSpcNodeType = "Pts";
      }
    }

    const lvlNode = this.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
    if (lvlNode !== undefined) {
      lvl = parseInt(lvlNode) + 1;
    }

    let fontSize: number | undefined = undefined;
    if (pNode["a:r"] !== undefined) {
      const fontSizeStr = this.getFontSize(pNode["a:r"], textBodyNode, undefined, lvl, type, warpObj);
      if (fontSizeStr !== "inherit") {
        fontSize = parseInt(fontSizeStr, 10);
      }
    }

    // 检查 layout 和 master
    const isInLayoutOrMaster = (type !== "shape" && type !== "textBox");
    if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
      // check in layout
      if (idx !== undefined) {
        const laypPrNode = this.getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx.toString(), "p:txBody", "a:p", (lvl - 1).toString(), "a:pPr"]);

        if (spcBefNode === undefined) {
          spcBefNode = this.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
        }
        if (spcAftNode === undefined) {
          spcAftNode = this.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
        }
        if (lnSpcNode === undefined) {
          lnSpcNode = this.getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
          if (lnSpcNode === undefined) {
            lnSpcNode = this.getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
            if (lnSpcNode !== undefined) {
              lnSpcNodeType = "Pts";
            }
          }
        }
      }
    }

    if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
      // check in master
      const slideMasterTextStyles = warpObj["slideMasterTextStyles"];
      let dirLoc = "";
      const lvlStr = `a:lvl${lvl}pPr`;

      switch (type) {
        case "title":
        case "ctrTitle":
          dirLoc = "p:titleStyle";
          break;
        case "body":
        case "obj":
        case "dt":
        case "ftr":
        case "sldNum":
        case "textBox":
          dirLoc = "p:bodyStyle";
          break;
        case "shape":
        default:
          dirLoc = "p:otherStyle";
      }

      const inLvlNode = this.getTextByPathList(slideMasterTextStyles, [dirLoc, lvlStr]);
      if (inLvlNode !== undefined) {
        if (spcBefNode === undefined) {
          spcBefNode = this.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
        }
        if (spcAftNode === undefined) {
          spcAftNode = this.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
        }
        if (lnSpcNode === undefined) {
          lnSpcNode = this.getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
          if (lnSpcNode === undefined) {
            lnSpcNode = this.getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
            if (lnSpcNode !== undefined) {
              lnSpcNodeType = "Pts";
            }
          }
        }
      }
    }

    let spcBefor = 0;
    let spcAfter = 0;
    let spcLines = 0;
    let marginTopBottomStr = "";

    if (spcBefNode !== undefined) {
      spcBefor = parseInt(spcBefNode) / 100;
    }
    if (spcAftNode !== undefined) {
      spcAfter = parseInt(spcAftNode) / 100;
    }

    if (lnSpcNode !== undefined && fontSize !== undefined) {
      if (lnSpcNodeType === "Pts") {
        marginTopBottomStr += `padding-top: ${((parseInt(lnSpcNode) / 100) - fontSize)}px;`;
      } else {
        const fct = parseInt(lnSpcNode) / 100000;
        spcLines = fontSize * (fct - 1) - fontSize;
        const pTop = (fct > 1) ? spcLines : 0;
        const pBottom = (fct > 1) ? fontSize : 0;
        marginTopBottomStr += `padding-top: ${pTop}px;`;
        marginTopBottomStr += `padding-bottom: ${spcLines}px;`;
      }
    }

    if (spcBefNode !== undefined || lnSpcNode !== undefined) {
      marginTopBottomStr += `margin-top: ${(spcBefor - 1)}px;`;
    }
    if (spcAftNode !== undefined || lnSpcNode !== undefined) {
      marginTopBottomStr += `margin-bottom: ${spcAfter}px;`;
    }

    return marginTopBottomStr;
  }

  /** 获取Layout和Master节点 */
  private getLayoutAndMasterNode(pNode: any, type: string): { nodeLaout: any; nodeMaster: any } {
    // 简化实现
    return { nodeLaout: undefined, nodeMaster: undefined };
  }

  /** 获取字体颜色属性 */
  private getFontColorPr(node: any, pNode: any, lstStyle: any, pFontStyle: any, lvl: number, idx: number | undefined, type: string, warpObj: any): any[] {
    // 返回 [color, textShadow, colorType, highlightColor]
    let color = "";
    const rPrNode = node["a:rPr"];
    
    if (rPrNode && rPrNode["a:solidFill"]) {
      color = this.getSolidFill(rPrNode["a:solidFill"]);
    }

    return [color, "", color ? "solid" : "", ""];
  }

  /** 获取纯色填充值 */
  private getSolidFill(solidFillNode: any): string {
    const srgbClr = solidFillNode?.["a:srgbClr"];
    if (srgbClr) {
      const val = this.getAttr(srgbClr, "val");
      return val || "";
    }
    return "";
  }

  /** 获取XML属性 */
  private getAttr(node: any, attrName: string): string | undefined {
    if (!node || !node.attrs) {
      return undefined;
    }
    return node.attrs[attrName];
  }

  /** 通过路径列表获取文本 - 简化版 */
  private getTextByPathList(node: any, path: string[]): any {
    if (!node) return undefined;
    let current = node;
    for (const p of path) {
      if (current && current[p] !== undefined) {
        current = current[p];
      } else {
        return undefined;
      }
    }
    return current;
  }

  /** 渲染图片 */
  private renderImage(image: any): string {
    const style: string[] = [
      `position:absolute`,
      `left:${image.x || 0}px`,
      `top:${image.y || 0}px`,
      `width:${image.width || 0}px`,
      `height:${image.height || 0}px`,
    ];

    // 使用已经提取的图片数据（base64或blob URL）
    const src = image.data || '';

    return `<img class="slide-image" style="${style.join(';')}" src="${src}" alt="slide image" />`;
  }

  /** 渲染表格 */
  private renderTable(table: any): string {
    let html = `<table class="slide-table">`;

    if (table.rows) {
      table.rows.forEach((row: any) => {
        html += '<tr>';
        if (row.cells) {
          row.cells.forEach((cell: any) => {
            html += '<td>';
            if (cell.textBlocks) {
              cell.textBlocks.forEach((block: any) => {
                html += this.renderTextBlock(block);
              });
            }
            html += '</td>';
          });
        }
        html += '</tr>';
      });
    }

    html += '</table>';
    return html;
  }

  /** 渲染图表 */
  private renderChart(chart: any): string {
    return `<div class="slide-chart" data-chart-type="${chart.chartType || 'unknown'}">Chart placeholder</div>`;
  }

  /** 获取样式类名，如果不存在则创建 */
  private getStyleClass(styleText: string): string {
    if (this.styleTable[styleText]) {
      return this.styleTable[styleText].name;
    }
    const className = `_css_${Object.keys(this.styleTable).length + 1}`;
    this.styleTable[styleText] = {
      name: className,
      text: styleText,
    };
    return className;
  }

  /** 生成全局CSS */
  private generateGlobalCSS(): string {
    let cssText = '';
    for (const key in this.styleTable) {
      const style = this.styleTable[key];
      cssText += `.text-block.${style.name}{${style.text}}\n`;
    }
    return cssText;
  }

  /** HTML 转义 */
  private escapeHtml(text: string): string {
    // @ts-ignore
    if(typeof text !== 'string') return text?.toString?.() || text;
    const map: Record<string, string> = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;',
    };
    return text.replace(/[&<>"']/g, (m) => map[m]);
  }

  /** 渲染文本块 - 旧版兼容 */
  private renderTextBlock(block: any, shape?: any): string {
    // 此方法仅作为回退，实际使用genTextBody
    return "";
  }
}
