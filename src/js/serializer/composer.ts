/**
 * PPTX Composer 模块（流式构建 API）
 *
 * 参考 nodejs-pptx 的 Composer 模式，提供两种等价写法：
 *
 * 1. 回调式（fluent）：
 *    const composer = new PPTXComposer();
 *    composer.addSlide(slide => {
 *        slide.addText(t => t.value('Hello').x(100).y(50).fontSize(24).bold());
 *        slide.addShape({ shapeType: 'roundRect', x: 0, y: 0, width: 200, height: 100, fill: { color: '#4f46e5' } });
 *    });
 *    const data = await composer.save();
 *
 * 2. 对象式：直接传配置对象给 addText/addShape/addImage。
 *
 * Composer 内部维护一棵「演示文稿 JSON 树」，可通过 toJSON() 导出，
 * 再交给 jsonToPptx() 序列化为 PPTX 文件（save() 即此封装）。
 *
 * @module serializer/composer
 */

import { jsonToPptx } from './json-to-pptx';

/** 默认 16:9 幻灯片尺寸（px，与解析端 SLIDE_FACTOR 换算一致：12192000EMU x 6858000EMU） */
const DEFAULT_SLIDE_SIZE = { width: 1280, height: 720 };

/**
 * 创建作用于目标对象的流式 setter 集合
 * @param {Object} el - 实际承载属性的目标对象
 * @param {string[]} keys - 属性名列表
 * @returns {Object} 构建器（每个 key 对应一个 setter，返回构建器自身以支持链式调用）
 */
function makeFluent(el: any, keys: any) {
    const builder: any = {};
    for (const key of keys) {
        builder[key] = (value: any) => {
    el[key] = value === undefined ? true : value;
    return builder;
};
    }
    return builder;
}

/**
 * 应用回调或对象配置
 * @param {Object} element - 元素对象
 * @param {Function|Object} config - 回调函数或配置对象
 */
function applyConfig(element: any, config: any) {
    if (typeof config === 'function') {
        config(element);
    } else if (config && typeof config === 'object') {
        Object.assign(element, config);
    }
}

/**
 * 幻灯片构建器
 */
class SlideComposer {
    slide: any;
    constructor() {
        /** @type {{background: *, elements: Array}} */
        this.slide = { background: null, elements: [] };
    }

    /**
     * 设置背景色
     * @param {string} color - 颜色值
     * @returns {SlideComposer} this
     */
    background(color: any) {
        this.slide.background = color;
        return this;
    }

    /**
     * 添加文本框
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象
     * @returns {SlideComposer} this
     */
    addText(config: any) {
        const el: any = { type: 'text', x: 0, y: 0, width: 300, height: 60 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'align', 'valign',
                'fontSize', 'color', 'bold', 'italic', 'underline', 'fontFace', 'href', 'lang', 'name']);
            builder.value = (text: any) => { el.text = text; return builder; };
            builder.runs = (runs: any) => { el.runs = runs; return builder; };
            builder.paragraphs = (paragraphs: any) => { el.paragraphs = paragraphs; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }

    /**
     * 添加形状
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象
     * @returns {SlideComposer} this
     */
    addShape(config: any) {
        const el: any = { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 200, height: 120 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'rotation', 'name']);
            builder.shapeType = (type: any) => { el.shapeType = type; return builder; };
            builder.fill = (fill: any) => { el.fill = fill; return builder; };
            builder.line = (line: any) => { el.line = line; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }

    /**
     * 添加图片
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象
     * @returns {SlideComposer} this
     */
    addImage(config: any) {
        const el: any = { type: 'image', x: 0, y: 0, width: 300, height: 200 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'extension', 'href', 'name']);
            builder.data = (data: any) => { el.data = data; return builder; };
            builder.src = (src: any) => { el.src = src; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }

    /**
     * 添加原生图表（PowerPoint/WPS 可编辑的真图表）
     * 配置对象字段：
     *   chartType: 'barChart'|'lineChart'|'areaChart'|'pieChart'|'pie3DChart'|'scatterChart'
     *   title, legend, varyColors, categories:[],
     *   series: [{ name, values:[] }]（非散点） 或  [{ name, x:[], y:[] }]（散点）
     * @param {Function|Object} config - 回调（接收配置对象，可直接赋值字段）或配置对象
     * @returns {SlideComposer} this
     */
    addChart(config: any) {
        const el: any = { type: 'chart', chartType: 'barChart', x: 0, y: 0, width: 600, height: 400 };
        if (typeof config === 'function') {
            config(el);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }
}

/**
 * 演示文稿构建器（Composer）
 */
class PPTXComposer {
    presentation: any;
    constructor() {
        this.presentation = {
            metadata: {},
            slideSize: { ...DEFAULT_SLIDE_SIZE },
            slides: []
        };
    }

    /**
     * 设置幻灯片尺寸（px）。兼容两种写法：
     *   slideSize(1280, 720) 或 slideSize({ width: 1280, height: 720 })
     * @param {number|Object} width - 宽，或含 width/height 的对象
     * @param {number} [height] - 高
     * @returns {PPTXComposer} this
     */
    slideSize(width: any, height: any) {
        if (width && typeof width === 'object') {
            this.presentation.slideSize = {
                width: width.width,
                height: width.height
            };
        } else {
            this.presentation.slideSize = { width, height };
        }
        return this;
    }

    /**
     * 设置元数据（字段与 pptxToJson 返回的 metadata 互通）
     * @param {Object} metadata - 元数据
     * @returns {PPTXComposer} this
     */
    metadata(metadata: any) {
        this.presentation.metadata = { ...this.presentation.metadata, ...metadata };
        return this;
    }

    /** 元数据便捷方法：标题 */
    title(value: any) { return this.metadata({ title: value }); }

    /** 元数据便捷方法：作者 */
    author(value: any) { return this.metadata({ author: value }); }

    /** 元数据便捷方法：主题 */
    subject(value: any) { return this.metadata({ subject: value }); }

    /** 元数据便捷方法：关键词 */
    keywords(value: any) { return this.metadata({ keywords: value }); }

    /** 元数据便捷方法：描述 */
    description(value: any) { return this.metadata({ description: value }); }

    /**
     * 添加幻灯片
     * @param {Function|Object} config - 回调（接收 SlideComposer）或幻灯片配置对象
     * @returns {PPTXComposer} this
     */
    addSlide(config: any) {
        const slideComposer = new SlideComposer();
        if (typeof config === 'function') {
            config(slideComposer);
        } else if (config && typeof config === 'object') {
            applyConfig(slideComposer.slide, config);
        }
        this.presentation.slides.push(slideComposer.slide);
        return this;
    }

    /**
     * 导出演示文稿 JSON 树（可直接传给 jsonToPptx）
     * @returns {Object} 演示文稿 JSON
     */
    toJSON() {
        return JSON.parse(JSON.stringify(this.presentation));
    }

    /**
     * 序列化为 PPTX 文件数据
     * @param {Object} [options] - 序列化选项（同 jsonToPptx options）
     * @returns {Promise<Uint8Array>} PPTX 文件二进制数据
     */
    save(options: any) {
        return jsonToPptx(this.toJSON(), options);
    }
}

export { PPTXComposer, SlideComposer, DEFAULT_SLIDE_SIZE };
