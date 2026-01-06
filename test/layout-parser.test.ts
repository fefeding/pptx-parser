/**
 * Layout Parser 测试
 */

import { describe, it, expect } from 'vitest';
import { parseAllSlideLayouts, mergeBackgrounds } from '../src/core/layout-parser';

describe('Layout Parser', () => {
  describe('mergeBackgrounds', () => {
    it('should prefer slide background over layout and master', () => {
      const slideBg = { type: 'color' as const, value: '#ff0000' };
      const layoutBg = { type: 'color' as const, value: '#00ff00' };
      const masterBg = { type: 'color' as const, value: '#0000ff' };

      const result = mergeBackgrounds(slideBg, layoutBg, masterBg);
      expect(result.value).toBe('#ff0000');
    });

    it('should prefer layout background when slide has no background', () => {
      const slideBg = { type: 'none' as const };
      const layoutBg = { type: 'color' as const, value: '#00ff00' };
      const masterBg = { type: 'color' as const, value: '#0000ff' };

      const result = mergeBackgrounds(slideBg, layoutBg, masterBg);
      expect(result.value).toBe('#00ff00');
    });

    it('should use master background when both slide and layout have no background', () => {
      const slideBg = { type: 'none' as const };
      const layoutBg = { type: 'none' as const };
      const masterBg = { type: 'color' as const, value: '#0000ff' };

      const result = mergeBackgrounds(slideBg, layoutBg, masterBg);
      expect(result.value).toBe('#0000ff');
    });

    it('should return default white when no backgrounds are provided', () => {
      const result = mergeBackgrounds();
      expect(result.type).toBe('color');
      expect(result.value).toBe('#ffffff');
    });

    it('should handle image backgrounds', () => {
      const slideBg = { type: 'image' as const, value: 'image1.png', relId: 'rId1' };
      const layoutBg = { type: 'image' as const, value: 'image2.png', relId: 'rId2' };
      const masterBg = { type: 'color' as const, value: '#0000ff' };

      const result = mergeBackgrounds(slideBg, layoutBg, masterBg);
      expect(result.type).toBe('image');
      expect(result.value).toBe('image1.png');
      expect(result.relId).toBe('rId1');
    });
  });

  describe('parseAllSlideLayouts', () => {
    it.skip('should parse all slide layouts from PPTX file', async () => {
      // 需要实际的PPTX文件进行测试
      // TODO: 添加集成测试
    });
  });
});
