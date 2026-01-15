/**
 * PPTX 常量定义
 * 用于存储 PPTX 解析过程中使用的各种常量
 */

const PPTXConstants = {
    /**
 * 幻灯片缩放因子
 * 用于将 EMU 单位转换为像素单位
 * EMU (English Metric Unit) = 1/914400 英寸
 * 96 DPI / 914400 EMU per inch
 */
SLIDE_FACTOR: 96 / 914400,

    /**
 * 字体大小缩放因子
 */
FONT_SIZE_FACTOR: 4 / 3.2,

    /**
 * RTL 语言列表
 * 支持从右到左书写的语言
 */
RTL_LANGS: ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR", "ur-PK"],

    /**
 * 列表样式映射
 */
LIST_STYLE_MAP: {
    'bullet': 'disc',
    'numbered': 'decimal',
    'alphabetic': 'lower-alpha'
}
};

export { PPTXConstants };

// Also export to global scope for backward compatibility
// window.PPTXConstants = PPTXConstants; // Removed for ES modules
