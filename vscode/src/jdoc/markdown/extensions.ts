/**
 * Markdown 扩展语法：流程图（Mermaid）与数学公式（KaTeX）
 *
 * 解析阶段只产出语义容器（保留源码作降级显示），真正的渲染交给浏览器端的
 * `renderMarkdownExtras()`——Mermaid / KaTeX 都依赖 DOM，Node 端无法渲染。
 *
 * @module markdown/extensions
 */

import type { MarkedExtension, TokenizerAndRendererExtension, Tokens } from 'marked';

/** 代码块语言别名 -> 图表类型 */
const DIAGRAM_LANGS = new Set(['mermaid']);
/** 代码块语言别名 -> 数学块 */
const MATH_LANGS = new Set(['math', 'tex', 'latex', 'katex']);

/** 属性值转义（放进 data-tex / data-diagram 供浏览器端取用） */
function attrValue(value: string): string {
    return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 文本转义（降级显示用） */
function escapeText(value: string): string {
    return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 数学块容器（$$…$$ / \[…\] / ```math 围栏） */
function mathBlockHtml(tex: string): string {
    const trimmed = tex.trim();
    return `<pre class="jdoc-math-block" data-tex="${attrValue(trimmed)}"><code>${escapeText(trimmed)}</code></pre>`;
}

/** 行内数学容器（$…$ / \(…\)） */
function mathInlineHtml(tex: string): string {
    const trimmed = tex.trim();
    return `<code class="jdoc-math-inline" data-tex="${attrValue(trimmed)}">${escapeText(trimmed)}</code>`;
}

/**
 * 生成 marked 扩展集合
 * @param options.enableMath - 是否启用数学语法
 * @param options.enableDiagrams - 是否启用 mermaid 图表语法
 */
export function markdownExtensions(
    options: { enableMath?: boolean; enableDiagrams?: boolean } = {}
): MarkedExtension {
    const useMath = options.enableMath !== false;
    const useDiagrams = options.enableDiagrams !== false;

    // ---- 代码块：```mermaid / ```math ----
    const fenced: TokenizerAndRendererExtension = {
        name: 'jdocDiagramCode',
        level: 'block',
        tokenizer(src: string): Tokens.Generic | undefined {
            if (!useDiagrams && !useMath) return undefined;
            const m = /^(?:```|~~~)([a-zA-Z0-9_-]*)[ \t]*\n([\s\S]*?)\n?(?:```|~~~)[ \t]*(?:\n|$)/.exec(src);
            if (!m) return undefined;
            const lang = (m[1] || '').toLowerCase();
            if (useDiagrams && DIAGRAM_LANGS.has(lang)) {
                return { type: 'jdocDiagramCode', raw: m[0], lang, text: m[2] || '' } as unknown as Tokens.Generic;
            }
            if (useMath && MATH_LANGS.has(lang)) {
                // 与图表共用同一 token 名：marked 的 renderer 按 token 名查找，
                // 返回未注册的类型会抛 "Token with ... type was not found"。
                return { type: 'jdocDiagramCode', raw: m[0], lang, text: m[2] || '' } as unknown as Tokens.Generic;
            }
            return undefined;
        },
        renderer(token: Tokens.Generic): string {
            const generic = token as unknown as { lang?: string; text?: string };
            const text = generic.text || '';
            if (generic.lang && DIAGRAM_LANGS.has(generic.lang)) {
                return `<pre class="jdoc-mermaid" data-diagram="mermaid"><code class="language-mermaid">${escapeText(text)}</code></pre>\n`;
            }
            return `${mathBlockHtml(text)}\n`;
        }
    };

    // ---- 块级公式：$$…$$ 与 \[…\]（独占若干行） ----
    const mathBlock: TokenizerAndRendererExtension = {
        name: 'jdocMathBlock',
        level: 'block',
        start(src: string): number | undefined {
            if (!useMath) return undefined;
            const candidates = [src.indexOf('$$'), src.indexOf('\\[')].filter((i) => i >= 0);
            return candidates.length ? Math.min(...candidates) : undefined;
        },
        tokenizer(src: string): Tokens.Generic | undefined {
            if (!useMath) return undefined;
            const dollar = /^\$\$([\s\S]+?)\$\$[ \t]*(?:\n|$)/.exec(src);
            if (dollar) return { type: 'jdocMathBlock', raw: dollar[0], text: dollar[1] } as unknown as Tokens.Generic;
            const bracket = /^\\\[([\s\S]+?)\\][ \t]*(?:\n|$)/.exec(src);
            if (bracket) return { type: 'jdocMathBlock', raw: bracket[0], text: bracket[1] } as unknown as Tokens.Generic;
            return undefined;
        },
        renderer(token: Tokens.Generic): string {
            return `${mathBlockHtml((token as unknown as { text?: string }).text || '')}\n`;
        }
    };

    // ---- 行内公式：$…$ 与 \(…\) ----
    const mathInline: TokenizerAndRendererExtension = {
        name: 'jdocMathInline',
        level: 'inline',
        start(src: string): number | undefined {
            if (!useMath) return undefined;
            const candidates = [src.search(/\$(?!\s)/), src.search(/\\\(/)]
                .map((i) => (i >= 0 ? i : -1))
                .filter((i) => i >= 0);
            return candidates.length ? Math.min(...candidates) : undefined;
        },
        tokenizer(src: string): Tokens.Generic | undefined {
            if (!useMath) return undefined;
            // 行内 $ 前后不能紧跟空白，避开 "价格 $5 和 $6" 这类普通用法
            const dollar = /^\$(?!\s)([^$\n]+?)(?<!\s)\$/.exec(src);
            if (dollar) return { type: 'jdocMathInline', raw: dollar[0], text: dollar[1] } as unknown as Tokens.Generic;
            const paren = /^\\\(([^)\n]+?)\\\)/.exec(src);
            if (paren) return { type: 'jdocMathInline', raw: paren[0], text: paren[1] } as unknown as Tokens.Generic;
            return undefined;
        },
        renderer(token: Tokens.Generic): string {
            return mathInlineHtml((token as unknown as { text?: string }).text || '');
        }
    };

    return { extensions: [fenced, mathBlock, mathInline] };
}
