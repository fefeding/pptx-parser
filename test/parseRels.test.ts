import { describe, it, expect } from 'vitest';
import { parseRels } from '../src/utils';

describe('parseRels - 关系文件解析测试', () => {
  it('应该正确解析标准的关系文件', () => {
    const relsXml = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.jpg"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

    const result = parseRels(relsXml);

    expect(result['rId1']).toEqual({
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: '../media/image1.png'
    });

    expect(result['rId2']).toEqual({
      id: 'rId2',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: '../media/image2.jpg'
    });

    expect(result['rId3']).toEqual({
      id: 'rId3',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
      target: '../slideLayouts/slideLayout1.xml'
    });
  });

  it('应该处理空的关系文件', () => {
    const relsXml = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const result = parseRels(relsXml);
    expect(result).toEqual({});
  });

  it('应该处理缺少属性的关系', () => {
    const relsXml = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="../media/image1.png"/>
</Relationships>`;

    const result = parseRels(relsXml);

    expect(result['rId1']).toEqual({
      id: 'rId1',
      type: '',
      target: '../media/image1.png'
    });
  });

  it('应该忽略没有Id的关系', () => {
    const relsXml = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="test" Target="test.xml"/>
  <Relationship Id="rId1" Target="real.xml"/>
</Relationships>`;

    const result = parseRels(relsXml);

    expect(result['rId1']).toBeDefined();
    expect(Object.keys(result)).toHaveLength(1);
  });
});
