/**
 * 单位转换工具单元测试
 * 对齐PPTXjs的单位转换逻辑
 * 
 * 测试重点：
 * 1. EMU↔PX转换准确性（对齐PPTXjs的slideFactor = 96/914400）
 * 2. 字体单位转换（对齐PPTXjs的fontSizeFactor = 4/3.2）
 * 3. 双向转换一致性
 * 4. 边界情况处理
 */

import { describe, it, expect } from 'vitest';
import {
  emu2px,
  px2emu,
  fontUnits2px,
  pt2emu,
  emu2pt,
  px2pt,
  pt2px,
  percentToPx,
  distanceEmu,
  diagonalEmu,
  isValidEmu
} from '../src/utils/unit-converter';

describe('单位转换工具 - EMU↔PX转换', () => {
  describe('emu2px - EMU转像素', () => {
    it('应该使用PPTXjs的slideFactor进行转换', () => {
      // PPTXjs: parseInt(emu * slideFactor), slideFactor = 96 / 914400
      // 914400 EMU = 96 PX
      expect(emu2px(914400)).toBe(96);
    });

    it('应该正确转换常见尺寸', () => {
      expect(emu2px(0)).toBe(0);
      expect(emu2px(457200)).toBe(48); // 半英寸
      expect(emu2px(1828800)).toBe(192); // 2英寸
      expect(emu2px(2743200)).toBe(288); // 3英寸
    });

    it('应该处理大尺寸（完整幻灯片）', () => {
      // 标准幻灯片尺寸：9144000 EMU (10英寸宽)
      expect(emu2px(9144000)).toBe(960);
      // 标准幻灯片高度：6858000 EMU (7.5英寸高)
      expect(emu2px(6858000)).toBe(720);
    });

    it('应该处理无效输入', () => {
      expect(emu2px(NaN)).toBe(0);
      expect(emu2px(undefined as any)).toBe(0);
      expect(emu2px(null as any)).toBe(0);
      expect(emu2px('invalid' as any)).toBe(0);
    });
  });

  describe('px2emu - 像素转EMU', () => {
    it('应该是emu2px的逆操作', () => {
      // 96 PX = 914400 EMU
      expect(px2emu(96)).toBe(914400);
    });

    it('应该正确转换常见尺寸', () => {
      expect(px2emu(0)).toBe(0);
      expect(px2emu(48)).toBe(457200);
      expect(px2emu(192)).toBe(1828800);
      expect(px2emu(288)).toBe(2743200);
    });

    it('应该处理大尺寸（完整幻灯片）', () => {
      expect(px2emu(960)).toBe(9144000);
      expect(px2emu(720)).toBe(6858000);
    });

    it('应该双向转换保持一致性', () => {
      const testCases = [0, 96, 192, 480, 720, 960, 100];
      testCases.forEach(px => {
        const emu = px2emu(px);
        const convertedPx = emu2px(emu);
        expect(convertedPx).toBe(px);
      });
    });

    it('应该处理无效输入', () => {
      expect(px2emu(NaN)).toBe(0);
      expect(px2emu(undefined as any)).toBe(0);
      expect(px2emu(null as any)).toBe(0);
    });
  });
});

describe('单位转换工具 - 字体单位转换', () => {
  describe('fontUnits2px - 字体单位转像素', () => {
    it('应该使用PPTXjs的fontSizeFactor进行转换', () => {
      // PPTXjs: parseInt(sz) / 100 * fontSizeFactor, fontSizeFactor = 4 / 3.2
      // 2800 font units = 28 pt = 28 * 1.25 = 35 px
      expect(fontUnits2px(2800)).toBe(35);
    });

    it('应该正确转换常见字体大小', () => {
      expect(fontUnits2px(1000)).toBe(12.5); // 10pt
      expect(fontUnits2px(1200)).toBe(15); // 12pt
      expect(fontUnits2px(1400)).toBe(17.5); // 14pt
      expect(fontUnits2px(1800)).toBe(22.5); // 18pt
      expect(fontUnits2px(2000)).toBe(25); // 20pt
      expect(fontUnits2px(3200)).toBe(40); // 32pt
      expect(fontUnits2px(4400)).toBe(55); // 44pt
    });

    it('应该处理小数字体', () => {
      expect(fontUnits2px(800)).toBe(10); // 8pt
      expect(fontUnits2px(600)).toBe(7.5); // 6pt
    });

    it('应该处理无效输入', () => {
      expect(fontUnits2px(NaN)).toBe(0);
      expect(fontUnits2px(undefined as any)).toBe(0);
      expect(fontUnits2px(null as any)).toBe(0);
    });
  });
});

describe('单位转换工具 - PT单位转换', () => {
  describe('pt2emu - 磅转EMU', () => {
    it('应该使用正确的转换因子', () => {
      // 1 pt = 12700 EMU
      expect(pt2emu(1)).toBe(12700);
    });

    it('应该正确转换常见磅值', () => {
      expect(pt2emu(0)).toBe(0);
      expect(pt2emu(10)).toBe(127000);
      expect(pt2emu(12)).toBe(152400);
      expect(pt2emu(18)).toBe(228600);
      expect(pt2emu(24)).toBe(304800);
    });

    it('应该双向转换保持一致性', () => {
      const testCases = [1, 10, 18, 24, 32, 44];
      testCases.forEach(pt => {
        const emu = pt2emu(pt);
        const convertedPt = emu2pt(emu);
        expect(convertedPt).toBeCloseTo(pt, 2);
      });
    });
  });

  describe('emu2pt - EMU转磅', () => {
    it('应该正确转换到磅', () => {
      expect(emu2pt(12700)).toBe(1);
      expect(emu2pt(25400)).toBe(2);
      expect(emu2pt(63500)).toBe(5);
    });
  });

  describe('px2pt - 像素转磅', () => {
    it('应该在96DPI下转换', () => {
      // 96DPI: 1 px = 0.75 pt
      expect(px2pt(16)).toBe(12);
      expect(px2pt(24)).toBe(18);
      expect(px2pt(32)).toBe(24);
    });

    it('应该双向转换保持一致性', () => {
      const testCases = [12, 16, 18, 24, 32];
      testCases.forEach(pt => {
        const px = pt2px(pt);
        const convertedPt = px2pt(px);
        expect(convertedPt).toBeCloseTo(pt, 2);
      });
    });
  });

  describe('pt2px - 磅转像素', () => {
    it('应该在96DPI下转换', () => {
      // 96DPI: 1 pt = 1.333 px (4/3)
      expect(pt2px(12)).toBe(16);
      expect(pt2px(18)).toBe(24);
      expect(pt2px(24)).toBe(32);
    });
  });
});

describe('单位转换工具 - 百分比计算', () => {
  describe('percentToPx - 百分比转像素', () => {
    it('应该正确计算百分比', () => {
      expect(percentToPx(50, 100)).toBe(50);
      expect(percentToPx(100, 100)).toBe(100);
      expect(percentToPx(25, 200)).toBe(50);
      expect(percentToPx(75, 400)).toBe(300);
    });

    it('应该处理边界值', () => {
      expect(percentToPx(0, 100)).toBe(0);
      expect(percentToPx(100, 100)).toBe(100);
    });

    it('应该处理无效输入', () => {
      expect(percentToPx(NaN, 100)).toBe(0);
      expect(percentToPx(50, NaN)).toBe(0);
    });
  });
});

describe('单位转换工具 - 几何计算', () => {
  describe('distanceEmu - EMU距离计算', () => {
    it('应该正确计算两点之间的距离', () => {
      // 3-4-5直角三角形
      expect(distanceEmu(0, 0, 3000, 4000)).toBe(5000);
      expect(distanceEmu(1000, 2000, 4000, 6000)).toBe(3605.551275463989);
    });

    it('应该处理相同点（距离为0）', () => {
      expect(distanceEmu(100, 200, 100, 200)).toBe(0);
    });

    it('应该处理直线距离', () => {
      expect(distanceEmu(0, 0, 1000, 0)).toBe(1000);
      expect(distanceEmu(0, 0, 0, 2000)).toBe(2000);
    });
  });

  describe('diagonalEmu - 对角线计算', () => {
    it('应该正确计算矩形对角线', () => {
      // 3-4-5矩形
      expect(diagonalEmu(3000, 4000)).toBe(5000);
      // 正方形
      expect(diagonalEmu(1000, 1000)).toBe(1414.2135623730951);
    });

    it('应该处理零尺寸', () => {
      expect(diagonalEmu(0, 0)).toBe(0);
      expect(diagonalEmu(1000, 0)).toBe(1000);
      expect(diagonalEmu(0, 2000)).toBe(2000);
    });
  });
});

describe('单位转换工具 - 验证函数', () => {
  describe('isValidEmu - EMU值验证', () => {
    it('应该接受有效的EMU值', () => {
      expect(isValidEmu(0)).toBe(true);
      expect(isValidEmu(914400)).toBe(true);
      expect(isValidEmu(9144000)).toBe(true);
      expect(isValidEmu(1000000)).toBe(true);
    });

    it('应该拒绝负值', () => {
      expect(isValidEmu(-1)).toBe(false);
      expect(isValidEmu(-100)).toBe(false);
    });

    it('应该拒绝无效值', () => {
      expect(isValidEmu(NaN)).toBe(false);
      expect(isValidEmu(undefined as any)).toBe(false);
      expect(isValidEmu(null as any)).toBe(false);
      expect(isValidEmu('invalid' as any)).toBe(false);
    });

    it('应该拒绝过大的值', () => {
      expect(isValidEmu(5278761)).toBe(false);
      expect(isValidEmu(10000000)).toBe(false);
    });
  });
});

describe('单位转换工具 - 综合场景测试', () => {
  it('应该正确处理完整幻灯片尺寸转换', () => {
    // 标准幻灯片：960x720px
    const widthEmu = px2emu(960);
    const heightEmu = px2emu(720);
    
    expect(widthEmu).toBe(9144000);
    expect(heightEmu).toBe(6858000);
    
    // 反向转换
    expect(emu2px(widthEmu)).toBe(960);
    expect(emu2px(heightEmu)).toBe(720);
  });

  it('应该正确计算幻灯片对角线', () => {
    const widthEmu = 9144000;
    const heightEmu = 6858000;
    const diagonal = diagonalEmu(widthEmu, heightEmu);
    
    // 计算：sqrt(9144000^2 + 6858000^2)
    const expected = Math.sqrt(widthEmu ** 2 + heightEmu ** 2);
    expect(diagonal).toBeCloseTo(expected, 2);
    
    // 转换为像素
    const diagonalPx = emu2px(diagonal);
    expect(diagonalPx).toBeCloseTo(1200, 0); // 960x720的16:9矩形对角线
  });

  it('应该正确处理字体大小和位置的组合转换', () => {
    // 模拟一个文本框场景
    const fontSizePt = 18; // 18pt字体
    const fontSizePx = pt2px(fontSizePt);
    const boxWidthPx = 200;
    const boxHeightPx = 100;
    
    // 转换为EMU
    const fontSizeEmu = pt2emu(fontSizePt);
    const boxWidthEmu = px2emu(boxWidthPx);
    const boxHeightEmu = px2emu(boxHeightPx);
    
    expect(fontSizePx).toBe(24);
    expect(fontSizeEmu).toBe(228600);
    expect(boxWidthEmu).toBe(1905000);
    expect(boxHeightEmu).toBe(952500);
    
    // 反向转换验证
    expect(emu2px(boxWidthEmu)).toBe(boxWidthPx);
    expect(emu2px(boxHeightEmu)).toBe(boxHeightPx);
  });
});