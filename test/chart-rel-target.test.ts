import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { pptxToHtml } from '../src/index.ts';

/**
 * 回归用例：部分生成器把 rels 的 Target 写成绝对路径（如 "/ppt/charts/chart2.xml"）。
 * 旧逻辑会拼成 "ppt" + target → "ppt/ppt/charts/chart2.xml"，部件定位失败，
 * 图表只剩空占位 div 且 result.charts 为空（图表永远渲染不出来）。
 */
describe('图表关系解析', () => {
    it('绝对路径 Target（/ppt/charts/...）同样能提取图表数据', async () => {
        const buf = readFileSync('examples/t_user数据质量报告.pptx');
        const result = await pptxToHtml(new Uint8Array(buf).buffer, {
            slideMode: false,
            keyBoardShortCut: false
        } as never);

        expect(result.slides.length).toBeGreaterThan(0);
        // 该文件含 3 个图表（chart2/3/4.xml）
        expect((result.charts || []).length).toBe(3);
    });
});
