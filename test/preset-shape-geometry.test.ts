// 预设几何（preset shape）渲染：编辑器与预览端共用的几何公式
// 覆盖此前只有 CSS 近似的形状（月亮/太阳/云朵/圆环…）以及旧公式画错的几何。
import { describe, it, expect } from 'vitest';
import { presetShapePath } from '../src/index.ts';

/** 取出路径里所有数字（用于检查是否越界 / 数值合法） */
function numbersOf(d: string): number[] {
  return (d.match(/-?\d+(?:\.\d+)?/g) || []).map(Number);
}

/** 统计 A（椭圆弧）指令数量 */
function arcCount(d: string): number {
  return (d.match(/A/gi) || []).length;
}

describe('预设几何 presetShapePath', () => {
    it('月亮：轮廓不越出形状框（旧公式会画到框外成尖刺）', () => {
        const w = 160, h = 160;
        const geo = presetShapePath('moon', w, h, {})!;
        expect(geo).toBeTruthy();
        expect(geo.d).toContain('A');
        // 路径坐标必须落在 [0,w]×[0,h] 内（允许极小的浮点误差）
        for (const n of numbersOf(geo.d)) {
            expect(n).toBeGreaterThanOrEqual(-0.5);
            expect(n).toBeLessThanOrEqual(Math.max(w, h) + 0.5);
        }
    });

    it('太阳：中心圆是一条真正的闭合椭圆（旧公式退化成零面积）', () => {
        const geo = presetShapePath('sun', 160, 160, {})!;
        expect(geo).toBeTruthy();
        // 8 条光芒 + 1 个中心圆 = 9 段
        expect((geo.d.match(/z/g) || []).length).toBe(9);
        // 中心圆由两段同半径的半椭圆弧组成
        expect(arcCount(geo.d)).toBeGreaterThanOrEqual(2);
    });

    it('云朵：11 段椭圆弧构成', () => {
        const geo = presetShapePath('cloud', 200, 120, {})!;
        expect(geo).toBeTruthy();
        expect(arcCount(geo.d)).toBe(11);
    });

    it('圆环：evenodd 填充 + 内外两层椭圆（孔是圆，不是四角星）', () => {
        const geo = presetShapePath('donut', 160, 160, {})!;
        expect(geo).toBeTruthy();
        expect(geo.fillRule).toBe('evenodd');
        // 外圈 + 内孔 = 4 段弧
        expect(arcCount(geo.d)).toBe(4);
    });

    it('其它复杂轮廓都有可用路径（非空且数值合法）', () => {
        const types = ['cloud', 'heart', 'lightningBolt', 'cube', 'bevel', 'foldedCorner',
            'corner', 'diagStripe', 'gear6', 'gear9', 'frame', 'halfFrame', 'sun', 'moon', 'donut'];
        for (const t of types) {
            const geo = presetShapePath(t, 160, 160, {});
            expect(geo, `${t} 应有几何路径`).toBeTruthy();
            expect(geo!.d.length).toBeGreaterThan(10);
            for (const n of numbersOf(geo!.d)) {
                expect(Number.isFinite(n)).toBe(true);
            }
        }
    });

    it('调整值生效：月牙深度随 adj 变化', () => {
        const thin = presetShapePath('moon', 160, 160, { adj: 90000 })!.d;
        const thick = presetShapePath('moon', 160, 160, { adj: 10000 })!.d;
        expect(thin).not.toBe(thick);
    });
});
