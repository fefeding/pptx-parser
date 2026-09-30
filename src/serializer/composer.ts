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
import type { ZipOutputType } from './json-to-pptx';
import type { SerializerElement, SerializerSlide, TextRunSpec, ParagraphSpec } from './element-builders';

/** 默认 16:9 幻灯片尺寸（px，与解析端 SLIDE_FACTOR 换算一致：12192000EMU x 6858000EMU） */
const DEFAULT_SLIDE_SIZE = { width: 1280, height: 720 };

/** 流式 setter 集合：每个 key 对应一个 setter，返回构建器自身以支持链式调用 */
type FluentBuilder = { [key: string]: (value?: unknown) => FluentBuilder };

/** 元素配置：回调（接收流式构建器）或配置对象 */
type ElementConfig = ((builder: FluentBuilder) => void) | Record<string, unknown>;

/** 图表元素配置：回调（接收元素对象）或配置对象 */
type ChartConfig = ((el: SerializerElement) => void) | Record<string, unknown>;

/** 演示文稿 JSON 树 */
export interface ComposerPresentation {
    metadata: Record<string, unknown>;
    slideSize: { width: number; height: number };
    slides: SerializerSlide[];
}

/**
 * 创建作用于目标对象的流式 setter 集合
 * @param {Object} el - 实际承载属性的目标对象
 * @param {string[]} keys - 属性名列表
 * @returns {Object} 构建器（每个 key 对应一个 setter，返回构建器自身以支持链式调用）
 */
function makeFluent(el: SerializerElement, keys: string[]): FluentBuilder {
    const builder: FluentBuilder = {};
    const target = el as unknown as Record<string, unknown>;
    for (const key of keys) {
        builder[key] = (value?: unknown) => {
            target[key] = value === undefined ? true : value;
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
function applyConfig(target: object, config: unknown) {
    if (typeof config === 'function') {
        (config as (t: object) => void)(target);
    } else if (config && typeof config === 'object') {
        Object.assign(target, config);
    }
}

/**
 * 幻灯片构建器
 */
class SlideComposer {
    slide: SerializerSlide;
    constructor() {
        /** @type {{background: *, elements: Array}} */
        this.slide = { background: null, elements: [] };
    }

    /**
     * 设置背景色
     * @param {string} color - 颜色值
     * @returns {SlideComposer} this
     */
    background(color: string) {
        this.slide.background = color;
        return this;
    }

    /**
     * 添加文本框
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象
     * @returns {SlideComposer} this
     */
    addText(config: ElementConfig) {
        const el: SerializerElement = { type: 'text', x: 0, y: 0, width: 300, height: 60 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'align', 'valign',
                'fontSize', 'color', 'bold', 'italic', 'underline', 'fontFace', 'href', 'lang', 'name',
                'lineSpacing', 'spaceBefore', 'spaceAfter', 'indentLeft', 'indentRight', 'indent', 'bullet']);
            builder.value = (text: unknown) => { el.text = text as string; return builder; };
            builder.runs = (runs: unknown) => { el.runs = runs as TextRunSpec[]; return builder; };
            builder.paragraphs = (paragraphs: unknown) => { el.paragraphs = paragraphs as ParagraphSpec[]; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements!.push(el);
        return this;
    }

    /**
     * 添加形状
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象
     * @returns {SlideComposer} this
     */
    addShape(config: ElementConfig) {
        const el: SerializerElement = { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 200, height: 120 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'rotation', 'name']);
            builder.shapeType = (type: unknown) => { el.shapeType = type as string; return builder; };
            builder.fill = (fill: unknown) => { el.fill = fill as SerializerElement['fill']; return builder; };
            builder.line = (line: unknown) => { el.line = line as SerializerElement['line']; return builder; };
            builder.effects = (effects: unknown) => { el.effects = effects as SerializerElement['effects']; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements!.push(el);
        return this;
    }

    /**
     * 添加图片
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象
     * @returns {SlideComposer} this
     */
    addImage(config: ElementConfig) {
        const el: SerializerElement = { type: 'image', x: 0, y: 0, width: 300, height: 200 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'extension', 'href', 'name']);
            builder.data = (data: unknown) => { el.data = data as string; return builder; };
            builder.src = (src: unknown) => { el.src = src as string; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements!.push(el);
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
    addChart(config: ChartConfig) {
        const el: SerializerElement = { type: 'chart', chartType: 'barChart', x: 0, y: 0, width: 600, height: 400 };
        if (typeof config === 'function') {
            config(el);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements!.push(el);
        return this;
    }

    /**
     * 添加分组（组合多个子元素，对应 p:grpSp）
     * 子元素坐标约定：默认 children 的 x/y 为相对组左上角的局部坐标（OOXML 标准）。
     * 若子元素使用页绝对坐标，请设 childrenCoordinates:'page'，生成时会自动减 group 偏移做相对化。
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象
     * @returns {SlideComposer} this
     */
    addGroup(config: ElementConfig) {
        const el: SerializerElement = { type: 'group', x: 0, y: 0, width: 400, height: 300, children: [] };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'name']);
            builder.children = (children: unknown) => { el.children = children as SerializerElement[]; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements!.push(el);
        return this;
    }

    /**
     * 添加 SmartArt 图示（对应 p:graphicFrame + 原生 diagrams/* 部件）
     * @param {Function|Object} config - 回调（接收流式构建器）或配置对象；
     *   关键字段：diagramType('list'|'hierarchy'|'process'|'cycle'|'pyramid')、nodes([{text, children?}])
     * @returns {SlideComposer} this
     */
    addDiagram(config: any) {
        const el: SerializerElement = { type: 'diagram', diagramType: 'list', x: 0, y: 0, width: 400, height: 300, nodes: [] };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'name', 'diagramType']);
            builder.nodes = (nodes: unknown) => { el.nodes = nodes as SerializerElement['nodes']; return builder; };
            config(builder);
        } else {
            applyConfig(el, config);
        }
        this.slide.elements!.push(el);
        return this;
    }
}
class PPTXComposer {
    presentation: ComposerPresentation;
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
    slideSize(width: number | { width: number; height: number }, height?: number) {
        if (width && typeof width === 'object') {
            this.presentation.slideSize = {
                width: width.width,
                height: width.height
            };
        } else {
            this.presentation.slideSize = { width: width as number, height: height as number };
        }
        return this;
    }

    /**
     * 设置元数据（字段与 pptxToJson 返回的 metadata 互通）
     * @param {Object} metadata - 元数据
     * @returns {PPTXComposer} this
     */
    metadata(metadata: Record<string, unknown>) {
        this.presentation.metadata = { ...this.presentation.metadata, ...metadata };
        return this;
    }

    /** 元数据便捷方法：标题 */
    title(value: string) { return this.metadata({ title: value }); }

    /** 元数据便捷方法：作者 */
    author(value: string) { return this.metadata({ author: value }); }

    /** 元数据便捷方法：主题 */
    subject(value: string) { return this.metadata({ subject: value }); }

    /** 元数据便捷方法：关键词 */
    keywords(value: string) { return this.metadata({ keywords: value }); }

    /** 元数据便捷方法：描述 */
    description(value: string) { return this.metadata({ description: value }); }

    /**
     * 添加幻灯片
     * @param {Function|Object} config - 回调（接收 SlideComposer）或幻灯片配置对象
     * @returns {PPTXComposer} this
     */
    addSlide(config: ((slide: SlideComposer) => void) | SerializerSlide) {
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
    save(options: { outputType?: ZipOutputType } = {}) {
        return jsonToPptx(this.toJSON(), options);
    }
}

export { PPTXComposer, SlideComposer, DEFAULT_SLIDE_SIZE };
