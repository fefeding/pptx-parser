/**
 * PDF 中文字体处理
 *
 * pdf-lib 自带的 14 种标准字体不含 CJK 字形，要输出中文必须内嵌 TTF/OTF，
 * 而内嵌依赖 pdf-lib 的 peer 依赖 `@pdf-lib/fontkit`（未注册时 embedFont 会抛
 * "Unknown font format"）。这里负责两件事：
 *
 * 1. `registerPdfFontkit()`：按需注册 fontkit（仅在真的要嵌入自定义字体时加载）
 * 2. `detectSystemFonts()`：在 Node 端探测系统里的中文字体，省得调用方自己找字体文件
 *
 * @module pdf/fonts
 */

import { JdocError } from '../core/errors';
import type { PdfFontSet } from './types';

/** fontkit 模块的最小接口（避免直接依赖其类型声明） */
type FontkitModule = { create(buffer: Uint8Array, postscriptName?: string): unknown };

let fontkitPromise: Promise<FontkitModule> | undefined;

/**
 * 注册 pdf-lib 的自定义字体引擎
 *
 * 重复调用是幂等的；未安装 `@pdf-lib/fontkit` 时抛出带安装指引的错误。
 */
export async function registerPdfFontkit(): Promise<FontkitModule> {
    if (!fontkitPromise) {
        fontkitPromise = import(/* @vite-ignore */ '@pdf-lib/fontkit')
            .then((mod) => (mod as { default?: FontkitModule }).default ?? (mod as unknown as FontkitModule))
            .catch((error) => {
                fontkitPromise = undefined;
                throw new JdocError(
                    '嵌入自定义字体需要 @pdf-lib/fontkit，请先安装：npm i @pdf-lib/fontkit',
                    'pdf',
                    { cause: error }
                );
            });
    }
    return fontkitPromise;
}

/**
 * 各平台常见中文字体候选路径
 *
 * pdf-lib 的 embedFont 只接受字体字节，无法指定 TTC 内的某个字面，
 * 因此优先列出单字体文件（.ttf/.otf）；.ttc 候选排在末位，
 * 由探测时的真实嵌入验证兜底（失败自动换下一个，不会产出缺字形的坏 PDF）。
 * DejaVu 等西文字体不含 CJK 字形，刻意不列，避免探测到后静默丢字。
 */
const SYSTEM_FONT_CANDIDATES: Record<string, string[]> = {
    darwin: [
        '/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
        '/Library/Fonts/Arial Unicode.ttf',
        '/System/Library/Fonts/Supplemental/Songti.ttc',
        '/System/Library/Fonts/STHeiti Light.ttc',
        '/System/Library/Fonts/Hiragino Sans GB.ttc',
        '/System/Library/Fonts/PingFang.ttc'
    ],
    win32: [
        'C:\\Windows\\Fonts\\simhei.ttf',
        'C:\\Windows\\Fonts\\Deng.ttf',
        'C:\\Windows\\Fonts\\msyh.ttc',
        'C:\\Windows\\Fonts\\simsun.ttc'
    ],
    linux: [
        '/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/arphic/uming.ttc'
    ]
};

/** 系统字体探测选项 */
export interface DetectSystemFontsOptions {
    /** 追加的候选文件路径（放在平台默认候选之前） */
    candidates?: string[];
    /** 是否包含 Windows 的用户字体目录，默认 true */
    includeUserFonts?: boolean;
}

/** 按当前平台取候选列表 */
function defaultCandidates(options: DetectSystemFontsOptions): string[] {
    const platform = (globalThis as { process?: { platform?: string } }).process?.platform || '';
    const list: string[] = [];
    if (platform === 'darwin') list.push(...(SYSTEM_FONT_CANDIDATES.darwin || []));
    else if (platform === 'win32') {
        list.push(...(SYSTEM_FONT_CANDIDATES.win32 || []));
        if (options.includeUserFonts !== false) list.push(`${process.env.LOCALAPPDATA || 'C:\\Users\\Default'}\\Microsoft\\Windows\\Fonts\\msyh.ttc`);
    } else list.push(...(SYSTEM_FONT_CANDIDATES.linux || []));

    return [...(options.candidates || []), ...list];
}

/** 探测结果缓存：同一进程只读一次磁盘 */
const detectCache = new Map<string, PdfFontSet | undefined>();

/**
 * 探测系统中可用的中文字体（仅 Node 端有效）
 *
 * 逐个候选尝试真实嵌入一次（fontkit 是否认得、pdf-lib 是否能子集化），
 * 首个成功的即返回，因此不会出现「探测成功但生成时报错」的假阳性。
 *
 * @param options - 候选路径覆盖等
 * @returns 字体字节集合；未找到可用中文字体时返回 undefined
 *
 * @example
 * ```js
 * const bytes = await jsonToPdf(spec, { fonts: await detectSystemFonts() });
 * ```
 */
export async function detectSystemFonts(options: DetectSystemFontsOptions = {}): Promise<PdfFontSet | undefined> {
    if (!(globalThis as { process?: { versions?: { node?: string } } }).process?.versions?.node) {
        throw new JdocError('系统字体探测仅支持 Node.js（浏览器端请自行传入 fonts）', 'pdf');
    }

    const key = JSON.stringify([options.candidates || [], options.includeUserFonts !== false]);
    if (detectCache.has(key)) return detectCache.get(key);

    const fs = (await import(/* @vite-ignore */ 'node:fs/promises')) as typeof import('node:fs/promises');
    const result = await findFirstUsableFont(defaultCandidates(options), fs);
    detectCache.set(key, result);
    return result;
}

/** 逐个候选做真实嵌入验证，返回首个可用字体的字节 */
async function findFirstUsableFont(
    candidates: string[],
    fs: typeof import('node:fs/promises')
): Promise<PdfFontSet | undefined> {
    const fontkit = await registerPdfFontkit();
    const { PDFDocument } = await import('pdf-lib');

    for (const path of candidates) {
        let bytes: Uint8Array;
        try {
            const stat = await fs.stat(path);
            if (!stat.isFile() || stat.size === 0) continue;
            bytes = new Uint8Array(await fs.readFile(path));
        } catch {
            continue;
        }

        try {
            const probe = await PDFDocument.create();
            probe.registerFontkit(fontkit as never);
            // 用与生产一致的 subset: true 探测：TTC 集合无法被 fontkit 子集化
            // （createSubset is not a function），只有子集化能成功才说明该候选真正可用。
            await probe.embedFont(bytes, { subset: true });
            // bold/italic 不单独探测：generate.ts 的 buildFonts 会自动回落到 regular
            return { regular: bytes };
        } catch {
            // 当前字体不可嵌入（如 TTC 集合），换下一个候选
        }
    }
    return undefined;
}
