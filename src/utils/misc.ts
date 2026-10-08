/**
 * 通用小工具（无依赖、无副作用）。
 *
 * 从编辑器 examples/editor/src/util.js 下沉的纯函数部分（uid/clamp/round/clone），
 * 供库内其它模块（如 editor 内核）复用，避免每个消费者各写一份。
 */

/** 生成近似唯一的 id（非密码学用途，仅用于文档内元素标识） */
export const uid = (p = 'e') => `${p}_${Math.random().toString(36).slice(2, 9)}${(Date.now() % 46656).toString(36)}`;

export const clamp = (v: number, a: number, b: number) => Math.min(b, Math.max(a, v));

export const round = (v: number, d = 2) => {
    const m = 10 ** d;
    return Math.round(v * m) / m;
};

/** 深拷贝（JSON 安全：文档模型只含 JSON 可序列化数据） */
export function clone<T>(o: T): T {
    return o == null ? o : JSON.parse(JSON.stringify(o)) as T;
}
