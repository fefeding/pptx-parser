import { describe, it, expect } from 'vitest';
import { jsonToPptx } from '../src/index';

/**
 * 高级 OOXML 能力生成端校验
 *
 * 覆盖此前「规范支持但生成端缺失 / 标完成实未达标」的能力：
 * 连接线 cxnSp、自定义几何 custGeom、3D（scene3d/sp3d）、字段 a:fld、
 * 文本框分栏与自动适配、艺术字变形、单元格合并标记、altText、
 * 图表次坐标轴/趋势线/轴标题、动画 presetClass 与触发、切换 advTm。
 */

async function slideXml(pres: any): Promise<string> {
    const buf: any = await jsonToPptx(pres);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    return await zip.file('ppt/slides/slide1.xml')!.async('text');
}

function oneSlide(elements: any[], slideExtra: any = {}) {
    return { slides: [{ background: '#ffffff', elements, ...slideExtra }] };
}

describe('高级 OOXML 生成能力', () => {
    it('连接线（p:cxnSp）：按起止点推导包围盒与方向', async () => {
        const xml = await slideXml(oneSlide([
            {
                type: 'connector', shapeType: 'bentConnector3',
                x: 100, y: 100, width: 200, height: 150,
                start: { x: 300, y: 250 }, end: { x: 100, y: 100 },
                line: { color: '#ff0000', width: 2, dashType: 'dash' }
            }
        ]));

        expect(xml).toContain('p:cxnSp');
        expect(xml).toContain('p:nvCxnSpPr');
        expect(xml).toContain('p:cNvCxnSpPr');
        // 起止点推导：off 取最小点，ext 取绝对值差
        expect(xml).toContain('prst="bentConnector3"');
        // 终点在起点左上 → 两个方向都翻转
        expect(xml).toContain('flipH="1"');
        expect(xml).toContain('flipV="1"');
        // 线型
        expect(xml).toContain('a:prstDash val="dash"');
        expect(xml).toContain('a:solidFill');
    });

    it('连接线不允许非法几何被回退：straightConnector1 保留', async () => {
        const xml = await slideXml(oneSlide([
            { type: 'connector', x: 0, y: 0, width: 100, height: 100, start: { x: 0, y: 0 }, end: { x: 100, y: 100 } }
        ]));
        // 缺省几何应为 straightConnector1，而非被白名单回退成 rect
        expect(xml).toContain('prst="straightConnector1"');
        expect(xml).not.toContain('prst="rect"');
    });

    it('自定义几何（a:custGeom）：优先于 prstGeom 并输出 pathLst', async () => {
        const xml = await slideXml(oneSlide([
            {
                type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 100, height: 100,
                custGeom: {
                    paths: [{
                        commands: [
                            { type: 'moveTo', x: 0, y: 0 },
                            { type: 'lnTo', x: 1, y: 0 },
                            { type: 'lnTo', x: 1, y: 1 },
                            { type: 'close' }
                        ]
                    }]
                }
            }
        ]));

        expect(xml).toContain('a:custGeom');
        expect(xml).toContain('a:pathLst');
        expect(xml).toContain('a:moveTo');
        expect(xml).toContain('a:lnTo');
        expect(xml).toContain('a:close');
        // 归一化坐标（0~1）应被放大到坐标空间
        expect(xml).toContain('x="100000"');
        // 有 custGeom 时不应再输出 prstGeom
        expect(xml).not.toContain('a:prstGeom');
    });

    it('三维效果：a:scene3d 与 a:sp3d 按规范顺序输出（effectLst 之后）', async () => {
        const xml = await slideXml(oneSlide([
            {
                type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 100, height: 100,
                fill: { color: '#336699' },
                threeD: {
                    shape: { extrusionHeight: 6, contourWidth: 1, material: 'plastic', bevelTop: { width: 4, height: 4 } },
                    scene: { camera: 'orthographicFront', fov: 30, lightRig: 'threePt', rotX: 10 }
                }
            }
        ]));

        expect(xml).toContain('a:sp3d');
        expect(xml).toContain('extrusionH=');
        expect(xml).toContain('prstMaterial="plastic"');
        expect(xml).toContain('a:bevelT');
        expect(xml).toContain('a:scene3d');
        expect(xml).toContain('a:camera');
        expect(xml).toContain('a:lightRig');
        // scene3d 必须排在 shape 填充/描边之后
        expect(xml.indexOf('a:scene3d')).toBeGreaterThan(xml.indexOf('a:solidFill'));
    });

    it('字段（a:fld）：页码/日期不被固化为普通 run', async () => {
        const xml = await slideXml(oneSlide([
            {
                type: 'text', x: 0, y: 0, width: 200, height: 40,
                runs: [{ text: '3', field: 'slidenum' }]
            }
        ]));

        expect(xml).toContain('a:fld');
        expect(xml).toContain('type="slidenum"');
        // a:fld 需要 id 属性（GUID）
        expect(xml).toMatch(/<a:fld id="\{[0-9A-F-]+\}"/);
    });

    it('文本框：分栏 / 自动适配 / 艺术字变形', async () => {
        const xml = await slideXml(oneSlide([
            {
                type: 'text', x: 0, y: 0, width: 300, height: 100, text: 'a',
                numCol: 2, spcCol: 12, autofit: 'normal', fontScale: 85,
                prstTxWarp: 'textArchUp'
            }
        ]));

        expect(xml).toContain('numCol="2"');
        expect(xml).toContain('spcCol=');
        expect(xml).toContain('a:normAutofit');
        expect(xml).toContain('fontScale=');
        expect(xml).toContain('a:prstTxWarp');
        expect(xml).toContain('prst="textArchUp"');
    });

    it('单元格合并：被吞并单元格输出 hMerge / vMerge', async () => {
        const xml = await slideXml(oneSlide([
            {
                type: 'table', x: 0, y: 0, width: 300, height: 100,
                rows: [
                    {
                        cells: [
                            { text: 'A', rowSpan: 2 },
                            { text: 'B' }
                        ]
                    },
                    {
                        cells: [
                            { text: '', vMerge: true },
                            { text: 'C' }
                        ]
                    }
                ]
            }
        ]));

        expect(xml).toContain('rowSpan="2"');
        expect(xml).toContain('vMerge="1"');
    });

    it('altText：p:cNvPr@descr 输出', async () => {
        const xml = await slideXml(oneSlide([
            { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 50, height: 50, descr: '装饰性矩形，供屏幕阅读器跳过' }
        ]));
        expect(xml).toContain('descr="装饰性矩形，供屏幕阅读器跳过"');
    });

    it('图表：次坐标轴生成第二组轴与独立图表节点', async () => {
        const buf: any = await jsonToPptx(oneSlide([
            {
                type: 'chart', chartType: 'barChart', x: 0, y: 0, width: 400, height: 300,
                categories: ['Q1', 'Q2'],
                secondaryValueAxis: true,
                axisTitles: { value: '金额', secondaryValue: '增长率' },
                series: [
                    { name: '收入', values: [10, 20] },
                    { name: '增长', values: [5, 8], axis: 'secondary' }
                ]
            }
        ]));
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const chartXml = await zip.file('ppt/charts/chart1.xml')!.async('text');

        // 两个图表节点（主轴组 + 次轴组）
        expect(chartXml.match(/<c:barChart>/g)?.length).toBe(2);
        // 次轴对
        expect(chartXml).toContain('<c:axId val="211"/>');
        expect(chartXml).toContain('<c:axId val="212"/>');
        // 轴标题
        expect(chartXml).toContain('<c:axTitle>');
        expect(chartXml).toContain('金额');
        expect(chartXml).toContain('增长率');
    });

    it('图表：趋势线按 CT_Trendline 顺序输出', async () => {
        const buf: any = await jsonToPptx(oneSlide([
            {
                type: 'chart', chartType: 'lineChart', x: 0, y: 0, width: 400, height: 300,
                categories: ['a', 'b', 'c'],
                series: [{
                    name: 'S', values: [1, 2, 3],
                    trendlines: [{ type: 'linear', name: '线性', forward: 1, showEquation: true, showRSquared: true }]
                }]
            }
        ]));
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const chartXml = await zip.file('ppt/charts/chart1.xml')!.async('text');

        expect(chartXml).toContain('<c:trendline>');
        expect(chartXml).toContain('<c:trendlineType val="linear"/>');
        expect(chartXml).toContain('<c:dispEq val="1"/>');
        expect(chartXml).toContain('<c:dispRSqr val="1"/>');
        expect(chartXml).toContain('<c:forward val="1"/>');
    });

    it('动画：presetClass / presetId / 点击触发', async () => {
        const xml = await slideXml(oneSlide(
            [{ type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 50, height: 50 }],
            {
                animations: [
                    { target: 0, type: 'flyIn', duration: 0.5, presetClass: 'entr' },
                    { target: 0, type: 'fade', duration: 0.3, presetClass: 'exit', trigger: { type: 'onClick' } }
                ]
            }
        ));

        expect(xml).toContain('p:timing');
        expect(xml).toContain('presetClass="entr"');
        expect(xml).toContain('presetClass="exit"');
        // flyIn 的预设编号为 2，不应恒为 1
        expect(xml).toContain('presetId="2"');
        // 点击触发：delay=indefinite
        expect(xml).toContain('delay="indefinite"');
        // preset 透传真实名称，不再收敛为四种
        expect(xml).toContain('preset="flyIn"');
    });

    it('切换：advTm / 禁止点击切换 / 方向', async () => {
        const xml = await slideXml(oneSlide([], {
            transition: { type: 'wipe', duration: 1000, advanceAfterTime: 5000, advanceOnClick: false, direction: 'l' }
        }));

        expect(xml).toContain('p:transition');
        expect(xml).toContain('advTm="5000"');
        expect(xml).toContain('advanceOnClick="0"');
        expect(xml).toContain('dir="l"');
    });

    it('公式元素（OMML）：mc:AlternateContent + a14:m + Fallback', async () => {
        const omml = '<m:oMathPara><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>';
        const xml = await slideXml(oneSlide([
            { type: 'math', x: 0, y: 0, width: 200, height: 60, omml, text: 'x^2' }
        ]));

        expect(xml).toContain('mc:AlternateContent');
        expect(xml).toContain('mc:Choice Requires="a14"');
        expect(xml).toContain('<a14:m>');
        // OMML 片段必须原样输出（未被 XML 转义）
        expect(xml).toContain('<m:oMathPara>');
        expect(xml).toContain('mc:Fallback');
    });

    it('OLE 对象：p:oleObj + 代理图 p:pic', async () => {
        const xml = await slideXml(oneSlide([
            {
                type: 'ole', x: 0, y: 0, width: 200, height: 150,
                progId: 'Excel.Sheet.12', data: 'AAAA', extension: 'xlsx'
            }
        ]));

        expect(xml).toContain('p:graphicFrame');
        expect(xml).toContain('p:oleObj');
        expect(xml).toContain('progId="Excel.Sheet.12"');
        // 必须内嵌 p:pic 作为显示代理
        expect(xml).toContain('p:nvPicPr');
    });
});
