/**
 * 形状分类常量模块
 * 定义所有PowerPoint预设形状的分类
 */

// 基础矩形类形状
export const RECT_SHAPES = [
    'rect',
    'flowChartProcess',
    'flowChartPredefinedProcess',
    'flowChartInternalStorage',
    'actionButtonBlank'
];

// 圆角矩形类
export const ROUND_RECT_SHAPES = [
    'roundRect',
    'round1Rect',
    'round2DiagRect',
    'round2SameRect',
    'flowChartAlternateProcess'
];

// 切角矩形类
export const SNIP_RECT_SHAPES = [
    'snip1Rect',
    'snip2DiagRect',
    'snip2SameRect',
    'flowChartPunchedCard'
];

// 流程图形状
export const FLOWCHART_SHAPES = [
    'flowChartCollate',
    'flowChartDocument',
    'flowChartMultidocument',
    'flowChartTerminator',
    'flowChartDecision',
    'flowChartSort',
    'flowChartMerge',
    'flowChartExtract',
    'flowChartInputOutput',
    'flowChartOffpageConnector',
    'flowChartOnlineStorage',
    'flowChartMagneticDisk',
    'flowChartMagneticTape',
    'flowChartOr',
    'flowChartSummingJunction',
    'flowChartDelay',
    'flowChartPreparation',
    'flowChartManualOperation',
    'flowChartManualInput',
    'flowChartCard',
    'flowChartPunchedTape',
    'flowChartConnector',
    'flowChartOffpageConnector',
    'flowChartDisplay',
    'flowChartLoopLimit',
    'flowChartStoredData'
];

// 按钮类
export const ACTION_BUTTONS = [
    'actionButtonBackPrevious',
    'actionButtonBeginning',
    'actionButtonDocument',
    'actionButtonEnd',
    'actionButtonForwardNext',
    'actionButtonHelp',
    'actionButtonHome',
    'actionButtonInformation',
    'actionButtonMovie',
    'actionButtonReturn',
    'actionButtonSound'
];

// 基础几何形状
export const BASIC_SHAPES = [
    'ellipse',
    'triangle',
    'diamond',
    'pentagon',
    'hexagon',
    'heptagon',
    'octagon',
    'decagon',
    'dodecagon',
    'pie',
    'pieWedge',
    'chord',
    'sector',
    'arc',
    'blockArc'
];

// 星形
export const STAR_SHAPES = [
    'star4',
    'star5',
    'star6',
    'star7',
    'star8',
    'star10',
    'star12',
    'star16',
    'star24',
    'star32'
];

// 箭头类
export const ARROW_SHAPES = [
    'bentArrow',
    'bentUpArrow',
    'curvedDownArrow',
    'curvedLeftArrow',
    'curvedRightArrow',
    'curvedUpArrow',
    'downArrow',
    'leftArrow',
    'leftRightArrow',
    'leftUpArrow',
    'notchedRightArrow',
    'quadArrow',
    'rightArrow',
    'stripedRightArrow',
    'upArrow',
    'upDownArrow',
    'uturnArrow',
    'leftCircularArrow',
    'leftRightCircularArrow',
    'swooshArrow'
];

// 标注/气泡类
export const CALLOUT_SHAPES = [
    'cloudCallout',
    'cloud',
    'rectangularCallout',
    'roundRectCallout',
    'ovalCallout',
    'wedgeRectCallout',
    'wedgeRoundRectCallout',
    'wedgeEllipseCallout',
    'borderCallout1',
    'borderCallout2',
    'borderCallout3',
    'accentCallout1',
    'accentCallout2',
    'accentCallout3',
    'callout1',
    'callout2',
    'callout3',
    'accentBorderCallout1',
    'accentBorderCallout2',
    'accentBorderCallout3'
];

// 括号类
export const BRACKET_SHAPES = [
    'bracePair',
    'bracketPair',
    'leftBrace',
    'leftBracket',
    'rightBrace',
    'rightBracket'
];

// 特殊形状
export const SPECIAL_SHAPES = [
    'irregularSeal1',
    'irregularSeal2',
    'gear6',
    'gear9',
    'moon',
    'corner',
    'diagStripe',
    'frame',
    'donut',
    'noSmoking',
    'plus',
    'minus',
    'multiply',
    'division',
    'equal',
    'notEqual'
];

/**
 * 获取形状所属分类
 * @param {string} shapeType - 形状类型
 * @returns {string} 分类名称
 */
export function getShapeCategory(shapeType) {
    if (RECT_SHAPES.includes(shapeType)) return 'rect';
    if (ROUND_RECT_SHAPES.includes(shapeType)) return 'roundRect';
    if (SNIP_RECT_SHAPES.includes(shapeType)) return 'snipRect';
    if (FLOWCHART_SHAPES.includes(shapeType)) return 'flowchart';
    if (ACTION_BUTTONS.includes(shapeType)) return 'actionButton';
    if (BASIC_SHAPES.includes(shapeType)) return 'basic';
    if (STAR_SHAPES.includes(shapeType)) return 'star';
    if (ARROW_SHAPES.includes(shapeType)) return 'arrow';
    if (CALLOUT_SHAPES.includes(shapeType)) return 'callout';
    if (BRACKET_SHAPES.includes(shapeType)) return 'bracket';
    if (SPECIAL_SHAPES.includes(shapeType)) return 'special';
    return 'unknown';
}

/**
 * 检查形状是否需要特殊处理
 * @param {string} shapeType - 形状类型
 * @returns {boolean}
 */
export function isComplexShape(shapeType) {
    return [
        ...ACTION_BUTTONS,
        'flowChartMultidocument',
        'actionButtonHelp',
        'actionButtonHome',
        'actionButtonInformation',
        'actionButtonMovie',
        'actionButtonReturn',
        'actionButtonSound'
    ].includes(shapeType);
}
