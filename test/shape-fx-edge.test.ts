import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx } from '../src/index.ts';

async function slideXml(data: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
}

function shapeDoc(effects: any) {
  return {
    slides: [{
      elements: [{
        type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 100, height: 60,
        fill: { color: '#FF0000' }, effects
      }]
    }]
  };
}

describe('P2d 形状特效边缘：reflection / softEdge / blur', () => {
  it('生成端将 reflection/softEdge/blur 合并进同一 a:effectLst（不与阴影/发光分离）', async () => {
    const buf = await jsonToPptx(shapeDoc({
      shadow: { color: '#000000', blur: 4, distance: 3, angle: 90 },
      reflection: { blur: 2, distance: 5, angle: 90, alpha: 50, scale: 100 },
      softEdge: { radius: 8 },
      blur: { radius: 4 }
    }));
    const xml = await slideXml(buf);
    expect(xml).toContain('<a:effectLst>');
    expect(xml).toContain('<a:reflection');
    expect(xml).toContain('<a:softEdge');
    expect(xml).toContain('<a:blur');
    // CT_ShapeProperties 顺序约束：整段只允许一个 a:effectLst
    const count = (xml.match(/<a:effectLst>/g) || []).length;
    expect(count).toBe(1);
    // reflection 关键属性正确写出
    expect(xml).toMatch(/<a:reflection[^>]*blurRad=/);
    expect(xml).toMatch(/<a:softEdge[^>]*rad=/);
    expect(xml).toMatch(/<a:blur[^>]*rad=/);
  });

  it('仅 reflection 时也能独立生成 effectLst（无阴影/发光也不丢）', async () => {
    const buf = await jsonToPptx(shapeDoc({ reflection: { distance: 10, alpha: 40 } }));
    const xml = await slideXml(buf);
    expect(xml).toContain('<a:reflection');
    expect(xml).toMatch(/<a:reflection[^>]*dist=/);
    expect(xml).toMatch(/<a:reflection[^>]*stA=/);
  });

  it('仅 blur/softEdge 时同样生成单个 effectLst', async () => {
    const buf = await jsonToPptx(shapeDoc({ softEdge: { radius: 6 }, blur: { radius: 3 } }));
    const xml = await slideXml(buf);
    expect(xml).toContain('<a:softEdge');
    expect(xml).toContain('<a:blur');
    const count = (xml.match(/<a:effectLst>/g) || []).length;
    expect(count).toBe(1);
  });
});
