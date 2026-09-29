/**
 * 核心数据结构类型定义
 *
 * 对应 tXml 解析器的「简化形态」：一个节点是以标签名为键的对象，
 * 其值为子节点 / 子节点数组 / 文本内容，并附带 attrs 属性集合。
 *
 * 迁移策略：先收紧「函数入参」以获得类型文档与调用检查，
 * 遍历内部仍用 any 游标（避免联合类型窄化带来的大量改动），返回值暂保持 any。
 */

/** XML 属性集合（tXml 会额外注入 order 记录节点顺序） */
export interface XmlAttrs {
    order?: number;
    [name: string]: any;
}

/** 子节点取值：单个节点 / 节点数组 / 文本 / 缺失（仅作文档说明） */
export type XmlValue = XmlNode | XmlNode[] | string | undefined;

/**
 * 单个 XML 节点（简化形态），例如 node["p:spPr"]["a:ln"]
 *
 * 索引签名暂用 any：tXml 产物的子标签形态多样，且代码中存在大量
 * 就地构造的对象字面量，过窄的联合类型会导致大量赋值不兼容。
 * 已获得的收益：attrs 具备类型、函数签名表达意图、后续可继续收敛。
 */
export interface XmlNode {
    attrs?: XmlAttrs;
    [tagName: string]: any;
}

/** 路径片段序列：标签名或数组下标，用于 getTextByPathList 等 */
export type XmlPath = (string | number)[];

/**
 * 解析共享上下文（在 shape / style / text 等模块间传递）
 * 注：仍保留索引签名，逐步收敛为具体字段类型
 */
export interface WarpObject {
    slideMasterTextStyles?: any;
    slideLayoutTextStyles?: any;
    slideMasterContent?: any;
    slideLayoutContent?: any;
    [key: string]: any;
}
