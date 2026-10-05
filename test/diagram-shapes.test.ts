import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { pptxToStandard } from '../src/index.js';

describe('SmartArt 图示缓存绘图形状提取', () => {
  it('从 drawingN.xml 提取树形布局形状（节点框 + 连接线 + 主题色解析）', async () => {
    const buf = readFileSync('/Users/jiamao/project/github/pptx-parser/examples/test-sample.pptx');
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const doc = await pptxToStandard(ab);
    let diagram: any = null;
    for (const slide of doc.slides) {
      const found = (slide.elements || []).find((e: any) => e.type === 'diagram' && e.shapes?.length);
      if (found) { diagram = found; break; }
    }
    expect(diagram).toBeTruthy();
    const shapes = diagram.shapes;
    // 节点框 + 连接线混合
    const nodes = shapes.filter((s: any) => !s.connector);
    const connectors = shapes.filter((s: any) => s.connector);
    expect(nodes.length).toBeGreaterThanOrEqual(4);
    expect(connectors.length).toBeGreaterThanOrEqual(3);
    // 根节点：圆角矩形 + 主题色填充 + 文字
    const root = nodes.find((s: any) => s.text === '根');
    expect(root).toBeTruthy();
    expect(root.prst).toBe('roundRect');
    expect(String(root.fill)).toMatch(/^#[0-9A-F]{6}$/i);
    expect(String(root.color)).toMatch(/^#[0-9A-F]{6}$/i); // scheme:lt1 已解析为具体色
    expect(root.fontSize).toBeGreaterThan(0);
    // 形状坐标应在图示框内
    for (const s of shapes) {
      expect(s.x).toBeGreaterThanOrEqual(-1);
      expect(s.y).toBeGreaterThanOrEqual(-1);
      expect(s.x + s.width).toBeLessThanOrEqual(diagram.width + 2);
    }
  });
});
