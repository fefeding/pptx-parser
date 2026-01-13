/**
 * PPTXjs文本处理工具 - TypeScript转译版
 * 对齐PPTXjs的文本处理逻辑
 */

/**
 * 文本对齐类型
 */
export enum TextAlign {
  LEFT = 'left',
  CENTER = 'center',
  RIGHT = 'right',
  JUSTIFY = 'justify',
  DISTRIBUTED = 'distributed',
}

/**
 * 垂直对齐类型
 */
export enum VerticalAlign {
  TOP = 'top',
  MIDDLE = 'middle',
  BOTTOM = 'bottom',
  JUSTIFY = 'justify',
  DISTRIBUTED = 'distributed',
}

/**
 * 文本样式结构
 */
export interface TextStyle {
  fontFace?: string;
  fontSize?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  baseline?: number; // 上标/下标，正值=上标，负值=下标
  textAlign?: TextAlign;
  textVerticalAlign?: VerticalAlign;
  lineSpacing?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  indent?: number;
  marginLeft?: number;
  marginRight?: number;
  textHighlight?: string;
  textShadow?: boolean;
}

/**
 * 文本段落结构
 */
export interface TextParagraph {
  text: string;
  styles?: TextStyle[];
  textAlign?: TextAlign;
  textVerticalAlign?: VerticalAlign;
  lineSpacing?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  indent?: number;
  marginLeft?: number;
  marginRight?: number;
}

/**
 * 文本运行结构
 */
export interface TextRun {
  text: string;
  style?: TextStyle;
}

/**
 * 解析文本属性 - 对齐PPTXjs的文本属性解析逻辑
 */
export function parseTextProps(textPropsNode: any): TextStyle {
  const style: TextStyle = {};

  // 1. 字体名称
  const latin = getTextByPathList(textPropsNode, ['a:latin', 'attrs', 'typeface']);
  const ea = getTextByPathList(textPropsNode, ['a:ea', 'attrs', 'typeface']);
  const cs = getTextByPathList(textPropsNode, ['a:cs', 'attrs', 'typeface']);
  
  if (latin) style.fontFace = latin;
  else if (ea) style.fontFace = ea;
  else if (cs) style.fontFace = cs;

  // 2. 字体大小
  const sz = getTextByPathList(textPropsNode, ['a:sz', 'attrs', 'val']);
  if (sz) {
    // PPTX字体单位是1/100点，转换为点
    style.fontSize = parseInt(sz) / 100;
  }

  // 3. 颜色
  const solidFill = textPropsNode['a:solidFill'];
  if (solidFill) {
    const { getColorValue } = require('./pptxjs-color-utils');
    const color = getColorValue(solidFill);
    if (color) {
      style.color = color;
    }
  }

  // 4. 粗体
  const b = getTextByPathList(textPropsNode, ['a:b', 'attrs', 'val']);
  style.bold = b !== '0'; // 默认为true

  // 5. 斜体
  const i = getTextByPathList(textPropsNode, ['a:i', 'attrs', 'val']);
  style.italic = i !== '0'; // 默认为true

  // 6. 下划线
  const u = getTextByPathList(textPropsNode, ['a:u', 'attrs', 'val']);
  style.underline = u !== 'none'; // 'none'表示无下划线

  // 7. 删除线
  const strike = getTextByPathList(textPropsNode, ['a:strike', 'attrs', 'val']);
  style.strike = strike !== 'noStrike'; // 'noStrike'表示无删除线

  // 8. 上标/下标
  const baseline = getTextByPathList(textPropsNode, ['a:baseline', 'attrs', 'val']);
  if (baseline) {
    style.baseline = parseInt(baseline);
  }

  return style;
}

/**
 * 解析段落属性 - 对齐PPTXjs的段落属性解析逻辑
 */
export function parseParagraphProps(paraPropsNode: any): Partial<TextStyle> {
  const props: Partial<TextStyle> = {};

  // 1. 对齐方式
  const algn = getTextByPathList(paraPropsNode, ['a:lnSpc', 'attrs', 'algn']);
  if (algn) {
    props.textAlign = algn as TextAlign;
  }

  // 2. 行间距
  const lnSpc = paraPropsNode['a:lnSpc'];
  if (lnSpc) {
    // 可能为固定值或百分比
    const spcPct = getTextByPathList(lnSpc, ['a:spcPct', 'attrs', 'val']);
    const spcPts = getTextByPathList(lnSpc, ['a:spcPts', 'attrs', 'val']);

    if (spcPct) {
      // 百分比行间距（单位为1/1000%）
      props.lineSpacing = parseInt(spcPct) / 1000;
    } else if (spcPts) {
      // 固定行间距（单位为1/100点）
      props.lineSpacing = parseInt(spcPts) / 100;
    }
  }

  // 3. 段前间距
  const spcBef = paraPropsNode['a:spcBef'];
  if (spcBef) {
    const spcPct = getTextByPathList(spcBef, ['a:spcPct', 'attrs', 'val']);
    const spcPts = getTextByPathList(spcBef, ['a:spcPts', 'attrs', 'val']);

    if (spcPct) {
      props.spacingBefore = parseInt(spcPct) / 1000;
    } else if (spcPts) {
      props.spacingBefore = parseInt(spcPts) / 100;
    }
  }

  // 4. 段后间距
  const spcAft = paraPropsNode['a:spcAft'];
  if (spcAft) {
    const spcPct = getTextByPathList(spcAft, ['a:spcPct', 'attrs', 'val']);
    const spcPts = getTextByPathList(spcAft, ['a:spcPts', 'attrs', 'val']);

    if (spcPct) {
      props.spacingAfter = parseInt(spcPct) / 1000;
    } else if (spcPts) {
      props.spacingAfter = parseInt(spcPts) / 100;
    }
  }

  // 5. 缩进
  const indent = getTextByPathList(paraPropsNode, ['a:marL', 'attrs', 'val']);
  if (indent) {
    // 单位为EMU，转换为点
    props.indent = parseInt(indent) / 914400;
  }

  // 6. 左边距
  const marL = getTextByPathList(paraPropsNode, ['a:marL', 'attrs', 'val']);
  if (marL) {
    props.marginLeft = parseInt(marL) / 914400;
  }

  // 7. 右边距
  const marR = getTextByPathList(paraPropsNode, ['a:marR', 'attrs', 'val']);
  if (marR) {
    props.marginRight = parseInt(marR) / 914400;
  }

  return props;
}

/**
 * 解析文本框内容 - 对齐PPTXjs的文本框解析逻辑
 */
export function parseTextBoxContent(txBodyNode: any): TextParagraph[] {
  const paragraphs: TextParagraph[] = [];

  if (!txBodyNode) {
    return paragraphs;
  }

  // 获取所有段落
  const paragraphsNode = txBodyNode['a:p'];
  const paragraphArray = Array.isArray(paragraphsNode) ? paragraphsNode : [paragraphsNode];

  for (const paragraph of paragraphArray) {
    const textParagraph: TextParagraph = {
      text: '',
      styles: [],
    };

    // 解析段落属性
    const pPr = paragraph['a:pPr'];
    if (pPr) {
      const paraProps = parseParagraphProps(pPr);
      Object.assign(textParagraph, paraProps);
    }

    // 解析文本运行
    const textRuns = paragraph['a:r'];
    const runArray = Array.isArray(textRuns) ? textRuns : [textRuns];

    const textParts: string[] = [];
    const styles: TextStyle[] = [];

    for (const run of runArray) {
      if (!run) continue;

      // 获取文本内容
      const t = getTextByPathList(run, ['a:t']);
      if (t) {
        textParts.push(t);
      }

      // 解析运行属性
      const rPr = run['a:rPr'];
      if (rPr) {
        const textStyle = parseTextProps(rPr);
        styles.push(textStyle);
      }
    }

    textParagraph.text = textParts.join('');
    if (styles.length > 0) {
      textParagraph.styles = styles;
    }

    paragraphs.push(textParagraph);
  }

  return paragraphs;
}

/**
 * 合并文本样式 - 对齐PPTXjs的样式继承逻辑
 */
export function mergeTextStyles(
  baseStyle: TextStyle,
  ...additionalStyles: (TextStyle | undefined)[]
): TextStyle {
  const merged: TextStyle = { ...baseStyle };

  for (const style of additionalStyles) {
    if (!style) continue;

    Object.assign(merged, style);
  }

  return merged;
}

/**
 * 生成文本CSS样式 - 对齐PPTXjs的CSS生成逻辑
 */
export function generateTextStyleCss(style: TextStyle): string {
  const cssStyles: string[] = [];

  // 1. 字体
  if (style.fontFace) {
    cssStyles.push(`font-family: "${style.fontFace}", Arial, sans-serif`);
  }

  // 2. 字体大小
  if (style.fontSize) {
    cssStyles.push(`font-size: ${style.fontSize}pt`);
  }

  // 3. 颜色
  if (style.color) {
    cssStyles.push(`color: ${style.color}`);
  }

  // 4. 粗体
  if (style.bold) {
    cssStyles.push('font-weight: bold');
  }

  // 5. 斜体
  if (style.italic) {
    cssStyles.push('font-style: italic');
  }

  // 6. 下划线
  if (style.underline) {
    cssStyles.push('text-decoration: underline');
  }

  // 7. 删除线
  if (style.strike) {
    if (cssStyles.includes('text-decoration: underline')) {
      cssStyles.push('text-decoration: underline line-through');
    } else {
      cssStyles.push('text-decoration: line-through');
    }
  }

  // 8. 上标/下标
  if (style.baseline) {
    if (style.baseline > 0) {
      cssStyles.push('vertical-align: super');
      cssStyles.push('font-size: smaller');
    } else if (style.baseline < 0) {
      cssStyles.push('vertical-align: sub');
      cssStyles.push('font-size: smaller');
    }
  }

  // 9. 对齐方式
  if (style.textAlign) {
    cssStyles.push(`text-align: ${style.textAlign}`);
  }

  // 10. 垂直对齐
  if (style.textVerticalAlign) {
    cssStyles.push(`vertical-align: ${style.textVerticalAlign}`);
  }

  // 11. 行间距
  if (style.lineSpacing) {
    if (style.lineSpacing < 1) {
      // 百分比行间距
      cssStyles.push(`line-height: ${Math.round(style.lineSpacing * 100)}%`);
    } else {
      // 固定行间距（点）
      cssStyles.push(`line-height: ${style.lineSpacing}pt`);
    }
  }

  // 12. 段前间距
  if (style.spacingBefore) {
    cssStyles.push(`margin-top: ${style.spacingBefore}pt`);
  }

  // 13. 段后间距
  if (style.spacingAfter) {
    cssStyles.push(`margin-bottom: ${style.spacingAfter}pt`);
  }

  // 14. 缩进
  if (style.indent) {
    cssStyles.push(`text-indent: ${style.indent}in`);
  }

  // 15. 左边距
  if (style.marginLeft) {
    cssStyles.push(`margin-left: ${style.marginLeft}in`);
  }

  // 16. 右边距
  if (style.marginRight) {
    cssStyles.push(`margin-right: ${style.marginRight}in`);
  }

  return cssStyles.join('; ');
}

/**
 * 生成文本段落HTML - 对齐PPTXjs的文本段落生成逻辑
 */
export function generateTextParagraphHtml(paragraph: TextParagraph): string {
  // 生成段落样式
  const paragraphStyle: Partial<TextStyle> = {
    textAlign: paragraph.textAlign,
    textVerticalAlign: paragraph.textVerticalAlign,
    lineSpacing: paragraph.lineSpacing,
    spacingBefore: paragraph.spacingBefore,
    spacingAfter: paragraph.spacingAfter,
    indent: paragraph.indent,
    marginLeft: paragraph.marginLeft,
    marginRight: paragraph.marginRight,
  };

  const paragraphCss = generateTextStyleCss(paragraphStyle);

  // 如果没有样式数组，生成简单段落
  if (!paragraph.styles || paragraph.styles.length === 0) {
    const css = paragraphCss ? ` style="${paragraphCss}"` : '';
    return `<p${css}>${paragraph.text}</p>`;
  }

  // 如果有样式数组，生成带样式的span
  let html = '';
  
  if (paragraphCss) {
    html += `<p style="${paragraphCss}">`;
  } else {
    html += '<p>';
  }

  // 分割文本并应用样式
  if (paragraph.styles.length === 1) {
    const runCss = generateTextStyleCss(paragraph.styles[0]);
    if (runCss) {
      html += `<span style="${runCss}">${paragraph.text}</span>`;
    } else {
      html += paragraph.text;
    }
  } else {
    // 如果有多个样式，尝试按字符分割
    // 简化处理：将整个文本包裹在第一个样式中
    const runCss = generateTextStyleCss(paragraph.styles[0]);
    if (runCss) {
      html += `<span style="${runCss}">${paragraph.text}</span>`;
    } else {
      html += paragraph.text;
    }
  }

  html += '</p>';

  return html;
}

/**
 * 生成文本框HTML - 对齐PPTXjs的文本框生成逻辑
 */
export function generateTextBoxHtml(paragraphs: TextParagraph[]): string {
  if (!paragraphs || paragraphs.length === 0) {
    return '';
  }

  return paragraphs
    .map(para => generateTextParagraphHtml(para))
    .join('');
}

/**
 * 处理文本换行 - 对齐PPTXjs的换行处理逻辑
 */
export function processTextLineBreaks(text: string): string {
  // 将\n、\r\n、\r转换为<br>
  return text.replace(/(\r\n|\r|\n)/g, '<br>');
}

/**
 * 获取默认文本样式 - 对齐PPTXjs的默认文本样式
 */
export function getDefaultTextStyle(): TextStyle {
  return {
    fontFace: 'Arial',
    fontSize: 18,
    color: '#000000',
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    baseline: 0,
    textAlign: TextAlign.LEFT,
    textVerticalAlign: VerticalAlign.TOP,
    lineSpacing: 1.0,
    spacingBefore: 0,
    spacingAfter: 0,
    indent: 0,
    marginLeft: 0,
    marginRight: 0,
  };
}

/**
 * 获取路径文本值 - 本地实现（避免循环依赖）
 */
export function getTextByPathList(obj: any, pathList: string[]): any {
  if (!obj || !pathList || pathList.length === 0) {
    return undefined;
  }

  let current = obj;
  for (const path of pathList) {
    if (current === undefined || current === null) {
      return undefined;
    }
    current = current[path];
  }

  return current;
}

/**
 * 导出函数
 */
