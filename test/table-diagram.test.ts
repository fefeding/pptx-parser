import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { extractSlideToStandard } from '../src/serializer/json-from-pptx';
import { PPTXXmlUtils } from '../src/utils/xml';

/** 含表格 + SmartArt 的幻灯片 OOXML */
const SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="2" name="Table 1"/>
          <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1600200"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr firstRow="1" bandRow="1"/>
              <a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid>
              <a:tr h="800100">
                <a:tc>
                  <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="1800" b="1"/><a:t>姓名</a:t></a:r></a:p></a:txBody>
                  <a:tcPr anchor="ctr"><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>分数</a:t></a:r></a:p></a:txBody>
                  <a:tcPr anchor="ctr"/>
                </a:tc>
              </a:tr>
              <a:tr h="800100">
                <a:tc>
                  <a:txBody><a:bodyPr/><a:lstStyle/>
                    <a:p><a:r><a:rPr lang="zh-CN"/><a:t>张三</a:t></a:r></a:p>
                    <a:p><a:r><a:rPr lang="zh-CN"/><a:t>补充</a:t></a:r></a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc gridSpan="1" rowSpan="1">
                  <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>95</a:t></a:r></a:p></a:txBody>
                  <a:tcPr/>
                </a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="3" name="Diagram 1"/>
          <p:cNvGraphicFramePr/>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm><a:off x="914400" y="3200400"/><a:ext cx="5486400" cy="2743200"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
            <dgm:rel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                     r:dm="rId7" r:lo="rId8" r:qs="rId9" r:cs="rId10"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;

/** SmartArt 数据部件（diagrams/data1.xml） */
const DIAGRAM_DATA_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <dgm:ptLst>
    <dgm:pt modelId="{1}" type="node">
      <dgm:t><a:p><a:r><a:rPr lang="zh-CN"/><a:t>第一步</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="{2}" type="node">
      <dgm:t><a:p><a:r><a:rPr lang="zh-CN"/><a:t>第二步</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
  </dgm:ptLst>
  <dgm:cxnLst/>
</dgm:dataModel>`;

async function buildSlideData() {
    const zip = new JSZip();
    zip.file('ppt/slides/slide1.xml', SLIDE_XML);
    zip.file('ppt/diagrams/data1.xml', DIAGRAM_DATA_XML);
    const slideContent = await PPTXXmlUtils.readXmlFile(zip, 'ppt/slides/slide1.xml');
    return {
        zip,
        slideData: {
            slideContent,
            slideResObj: { rId7: { type: 'diagramData', target: '../diagrams/data1.xml' } }
        }
    };
}

describe('表格与 SmartArt 提取（解析端语义化）', () => {
    it('提取表格的行列、文本、样式与合并', async () => {
        const { zip, slideData } = await buildSlideData();
        const slide = await extractSlideToStandard(slideData as any, zip as any);

        const table: any = slide.elements.find((e) => e.type === 'table');
        expect(table).toBeTruthy();
        // 914400EMU = 96px，5486400EMU = 576px
        expect(table.x).toBeCloseTo(96, 0);
        expect(table.width).toBeCloseTo(576, 0);
        expect(table.rows.length).toBe(2);

        // 表头：文本 + 底色 + 垂直居中 + 加粗
        const head = table.rows[0].cells;
        expect(head.map((c: any) => c.text)).toEqual(['姓名', '分数']);
        expect(head[0].fill).toBe('DDEEFF');
        expect(head[0].valign).toBe('middle');
        expect(head[0].bold).toBe(true);
        expect(head[0].fontSize).toBe(18);

        // 多段单元格保留 paragraphs
        const multi = table.rows[1].cells[0];
        expect(multi.paragraphs).toBeTruthy();
        expect(multi.paragraphs.map((p: any) => p.runs[0].text)).toEqual(['张三', '补充']);

        // 列宽 / 行高（2743200EMU = 288px，800100EMU = 84px）
        expect(table.colWidths.map((w: number) => Math.round(w))).toEqual([288, 288]);
        expect(table.rowHeights?.length).toBe(2);
        expect(table.rowHeights[0]).toBeCloseTo(84, 1);
    });

    it('提取 SmartArt（图示）的文本内容', async () => {
        const { zip, slideData } = await buildSlideData();
        const slide = await extractSlideToStandard(slideData as any, zip as any);

        const diagram: any = slide.elements.find((e) => e.type === 'diagram');
        expect(diagram).toBeTruthy();
        expect(diagram.texts).toEqual(['第一步', '第二步']);
        expect(diagram.dataPath).toBe('ppt/diagrams/data1.xml');
        expect(diagram.name).toBe('Diagram 1');
        // 语义层无法表达图示几何，保留原始节点兜底
        expect(diagram.__raw).toBeTruthy();
    });
});
