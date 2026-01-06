import { describe, it, expect } from 'vitest';
import { parseTheme, resolveSchemeColor } from '../src/core/theme-parser';

describe('Theme Parser', () => {
  describe('parseTheme', () => {
    it('should parse theme from XML', async () => {
      const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements>
          <a:clrScheme name="Office">
            <a:bg1>
              <a:srgbClr val="FFFFFF"/>
            </a:bg1>
            <a:tx1>
              <a:srgbClr val="000000"/>
            </a:tx1>
            <a:accent1>
              <a:srgbClr val="4472C4"/>
            </a:accent1>
            <a:accent2>
              <a:srgbClr val="ED7D31"/>
            </a:accent2>
            <a:accent3>
              <a:srgbClr val="A5A5A5"/>
            </a:accent3>
          </a:clrScheme>
        </a:themeElements>
      </a:theme>`;

      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();
      zip.file('ppt/theme/theme1.xml', themeXml);

      const theme = await parseTheme(zip);

      expect(theme).not.toBeNull();
      expect(theme?.colors).toBeDefined();
      expect(theme?.colors.bg1).toBe('#FFFFFF');
      expect(theme?.colors.tx1).toBe('#000000');
      expect(theme?.colors.accent1).toBe('#4472C4');
      expect(theme?.colors.accent2).toBe('#ED7D31');
      expect(theme?.colors.accent3).toBe('#A5A5A5');
    });

    it('should handle missing theme file', async () => {
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();

      const theme = await parseTheme(zip);

      expect(theme).toBeNull();
    });
  });

  describe('resolveSchemeColor', () => {
    it('should resolve scheme color without mapping', () => {
      const themeColors = {
        bg1: '#FFFFFF',
        tx1: '#000000',
        accent1: '#4472C4'
      };

      const color = resolveSchemeColor('accent1', themeColors);
      expect(color).toBe('#4472C4');
    });

    it('should resolve scheme color with mapping', () => {
      const themeColors = {
        bg1: '#FFFFFF',
        tx1: '#000000',
        dk1: '#000000',
        lt1: '#FFFFFF',
        accent1: '#4472C4'
      };

      const colorMap = {
        bg1: 'light'  // Map bg1 to lt1
      };

      const color = resolveSchemeColor('bg1', themeColors, colorMap);
      expect(color).toBe('#FFFFFF');  // Should return lt1 color
    });

    it('should return default color for unknown scheme', () => {
      const themeColors = {
        bg1: '#FFFFFF',
        tx1: '#000000'
      };

      const color = resolveSchemeColor('accent99', themeColors);
      expect(color).toBe('#ffffff');
    });
  });
});
