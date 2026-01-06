/**
 * 绘图解析器
 * 解析图表和 SmartArt 的 drawings 文件
 * 对齐 PPTXjs 的绘图解析能力
 */

import JSZip from 'jszip';
import { getFirstChildByTagNS, getAttrSafe, emu2px, log } from '../utils';
import { NS } from '../constants';
import type { RelsMap, ChartResult, DiagramResult, ChartSeries, ChartDataPoint, DiagramShape } from './types';

/**
 * 解析所有图表
 * @param zip JSZip对象
 * @returns 图表数组
 */
export async function parseAllCharts(zip: JSZip): Promise<ChartResult[]> {
  try {
    const chartFiles = Object.keys(zip.files)
      .filter(path => path.startsWith('ppt/charts/'))
      .filter(path => path.endsWith('.xml'))
      .filter(path => !path.includes('_rels'))
      .sort();

    log('info', `Found ${chartFiles.length} chart files`);

    const charts: ChartResult[] = [];

    for (const chartPath of chartFiles) {
      const chartXml = await zip.file(chartPath)?.async('string');
      if (!chartXml) {
        log('warn', `Failed to read chart: ${chartPath}`);
        continue;
      }

      const chartId = chartPath.match(/chart(\d+)\.xml/)?.[1] || '1';
      const relsMap = await parseChartRels(zip, `chart${chartId}`);

      const chart = parseChart(chartXml, chartId, relsMap);
      charts.push(chart);
    }

    return charts;
  } catch (error) {
    log('error', 'Failed to parse charts', error);
    return [];
  }
}

/**
 * 解析单个图表
 * @param chartXml 图表XML
 * @param chartId 图表ID
 * @param relsMap 关联关系映射表
 * @returns 图表解析结果
 */
function parseChart(chartXml: string, chartId: string, relsMap: RelsMap): ChartResult {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(chartXml, 'application/xml');
    const root = doc.documentElement;

    // 检测图表类型
    const chartType = detectChartType(root);

    // 解析标题
    const title = parseChartTitle(root);

    // 解析系列数据
    const series = parseChartSeries(root);

    // 解析分类数据
    const categories = parseChartCategories(root);

    // 解析轴标题
    const xTitle = parseAxisTitle(root, 'x');
    const yTitle = parseAxisTitle(root, 'y');

    // 解析图例和数据标签
    const showLegend = parseLegendVisibility(root);
    const showDataLabels = parseDataLabelsVisibility(root);

    log('info', `Parsed chart ${chartId}: type=${chartType}, series=${series.length}`);

    return {
      id: `chart_${chartId}`,
      chartType,
      title,
      series,
      categories,
      xTitle,
      yTitle,
      showLegend,
      showDataLabels,
      relsMap
    };
  } catch (error) {
    log('error', `Failed to parse chart ${chartId}`, error);
    return {
      id: `chart_${chartId}`,
      chartType: 'unknown',
      relsMap
    };
  }
}

/**
 * 检测图表类型
 */
function detectChartType(root: Element): 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown' {
  // 查找 plotArea
  const plotArea = getFirstChildByTagNS(root, 'plotArea', NS.c);
  if (!plotArea) return 'unknown';

  // 遍历 plotArea 的子元素查找图表类型
  const children = Array.from(plotArea.children);
  for (const child of children) {
    const tagName = child.tagName;

    if (tagName.includes('lineChart')) return 'lineChart';
    if (tagName.includes('barChart')) return 'barChart';
    if (tagName.includes('pieChart')) {
      if (tagName.includes('3D')) return 'pie3DChart';
      return 'pieChart';
    }
    if (tagName.includes('areaChart')) return 'areaChart';
    if (tagName.includes('scatterChart')) return 'scatterChart';
  }

  return 'unknown';
}

/**
 * 解析图表标题
 */
function parseChartTitle(root: Element): string | undefined {
  const title = getFirstChildByTagNS(root, 'title', NS.c);
  if (!title) return undefined;

  const tx = getFirstChildByTagNS(title, 'tx', NS.c);
  if (!tx) return undefined;

  const strRef = getFirstChildByTagNS(tx, 'strRef', NS.c);
  const strCache = getFirstChildByTagNS(tx, 'strCache', NS.c);

  if (strCache) {
    const pt = getFirstChildByTagNS(strCache, 'pt', NS.c);
    if (pt?.textContent) {
      return pt.textContent.trim();
    }
  }

  const rich = getFirstChildByTagNS(tx, 'rich', NS.c);
  if (rich) {
    const p = getFirstChildByTagNS(rich, 'p', NS.a);
    if (p) {
      const text = extractTextFromParagraph(p);
      if (text) return text;
    }
  }

  return undefined;
}

/**
 * 解析图表系列
 */
function parseChartSeries(root: Element): ChartSeries[] {
  const series: ChartSeries[] = [];

  const plotArea = getFirstChildByTagNS(root, 'plotArea', NS.c);
  if (!plotArea) return series;

  // 查找所有系列节点（lineChart/barChart等）
  const chartContainer = findChartContainer(plotArea);
  if (!chartContainer) return series;

  const serNodes = Array.from(chartContainer.children).filter(
    child => child.tagName.includes(':ser') || child.localName === 'ser'
  );

  for (const serNode of serNodes) {
    const idx = parseInt(getAttrSafe(serNode, 'idx', '0'), 10);
    const order = parseInt(getAttrSafe(serNode, 'order', '0'), 10);

    // 解析系列名称
    const name = parseSeriesName(serNode);

    // 解析数据点
    const points = parseSeriesPoints(serNode);

    series.push({
      name,
      idx,
      order,
      points
    });
  }

  return series;
}

/**
 * 查找图表容器
 */
function findChartContainer(plotArea: Element): Element | null {
  const children = Array.from(plotArea.children);
  for (const child of children) {
    const tagName = child.tagName;
    if (
      tagName.includes('lineChart') ||
      tagName.includes('barChart') ||
      tagName.includes('pieChart') ||
      tagName.includes('areaChart') ||
      tagName.includes('scatterChart')
    ) {
      return child;
    }
  }
  return null;
}

/**
 * 解析系列名称
 */
function parseSeriesName(serNode: Element): string | undefined {
  const tx = getFirstChildByTagNS(serNode, 'tx', NS.c);
  if (!tx) return undefined;

  const strRef = getFirstChildByTagNS(tx, 'strRef', NS.c);
  const strCache = getFirstChildByTagNS(tx, 'strCache', NS.c);

  if (strCache) {
    const pt = getFirstChildByTagNS(strCache, 'pt', NS.c);
    if (pt?.textContent) {
      return pt.textContent.trim();
    }
  }

  const v = getFirstChildByTagNS(tx, 'v', NS.c);
  if (v?.textContent) {
    return v.textContent.trim();
  }

  return undefined;
}

/**
 * 解析系列数据点
 */
function parseSeriesPoints(serNode: Element): ChartDataPoint[] {
  const points: ChartDataPoint[] = [];

  const numRef = getFirstChildByTagNS(serNode, 'val', NS.c);
  const numCache = getFirstChildByTagNS(serNode, 'numCache', NS.c);

  if (numCache) {
    const ptNodes = Array.from(numCache.children).filter(
      child => child.tagName.includes(':pt') || child.localName === 'pt'
    );

    for (const pt of ptNodes) {
      const idx = parseInt(getAttrSafe(pt, 'idx', '0'), 10);
      const v = getFirstChildByTagNS(pt, 'v', NS.c);
      const value = v?.textContent ? parseFloat(v.textContent) : undefined;

      points.push({ idx, value });
    }
  }

  return points;
}

/**
 * 解析图表分类
 */
function parseChartCategories(root: Element): string[] {
  const categories: string[] = [];

  const plotArea = getFirstChildByTagNS(root, 'plotArea', NS.c);
  if (!plotArea) return categories;

  const catAx = getFirstChildByTagNS(plotArea, 'catAx', NS.c);
  if (!catAx) return categories;

  const tx = getFirstChildByTagNS(catAx, 'tx', NS.c);
  if (!tx) return categories;

  const strRef = getFirstChildByTagNS(tx, 'strRef', NS.c);
  const strCache = getFirstChildByTagNS(tx, 'strCache', NS.c);

  if (strCache) {
    const ptNodes = Array.from(strCache.children).filter(
      child => child.tagName.includes(':pt') || child.localName === 'pt'
    );

    for (const pt of ptNodes) {
      const v = getFirstChildByTagNS(pt, 'v', NS.c);
      if (v?.textContent) {
        categories.push(v.textContent.trim());
      }
    }
  }

  return categories;
}

/**
 * 解析轴标题
 */
function parseAxisTitle(root: Element, axis: 'x' | 'y'): string | undefined {
  const plotArea = getFirstChildByTagNS(root, 'plotArea', NS.c);
  if (!plotArea) return undefined;

  const ax = axis === 'x'
    ? getFirstChildByTagNS(plotArea, 'catAx', NS.c)
    : getFirstChildByTagNS(plotArea, 'valAx', NS.c);

  if (!ax) return undefined;

  const title = getFirstChildByTagNS(ax, 'title', NS.c);
  if (!title) return undefined;

  const tx = getFirstChildByTagNS(title, 'tx', NS.c);
  if (!tx) return undefined;

  const rich = getFirstChildByTagNS(tx, 'rich', NS.c);
  if (rich) {
    const p = getFirstChildByTagNS(rich, 'p', NS.a);
    if (p) {
      const text = extractTextFromParagraph(p);
      if (text) return text;
    }
  }

  return undefined;
}

/**
 * 解析图例可见性
 */
function parseLegendVisibility(root: Element): boolean {
  const legend = getFirstChildByTagNS(root, 'legend', NS.c);
  if (!legend) return false;

  const deleted = legend.getAttribute('deleted');
  return deleted !== '1';
}

/**
 * 解析数据标签可见性
 */
function parseDataLabelsVisibility(root: Element): boolean {
  const plotArea = getFirstChildByTagNS(root, 'plotArea', NS.c);
  if (!plotArea) return false;

  const chartContainer = findChartContainer(plotArea);
  if (!chartContainer) return false;

  const dLbls = getFirstChildByTagNS(chartContainer, 'dLbls', NS.c);
  if (!dLbls) return false;

  const deleted = dLbls.getAttribute('deleted');
  return deleted !== '1';
}

/**
 * 从段落中提取文本
 */
function extractTextFromParagraph(p: Element): string | undefined {
  const rNodes = Array.from(p.children).filter(
    child => child.tagName.includes(':r') || child.localName === 'r'
  );

  const texts: string[] = [];
  for (const r of rNodes) {
    const t = getFirstChildByTagNS(r, 't', NS.a);
    if (t?.textContent) {
      texts.push(t.textContent);
    }
  }

  return texts.length > 0 ? texts.join('') : undefined;
}

/**
 * 解析所有 SmartArt/Diagram
 * @param zip JSZip对象
 * @returns Diagram 数组
 */
export async function parseAllDiagrams(zip: JSZip): Promise<DiagramResult[]> {
  try {
    const diagramFiles = Object.keys(zip.files)
      .filter(path => path.startsWith('ppt/diagrams/'))
      .filter(path => path.endsWith('.xml'))
      .filter(path => !path.includes('_rels'))
      .sort();

    log('info', `Found ${diagramFiles.length} diagram files`);

    const diagrams: DiagramResult[] = [];

    for (const diagramPath of diagramFiles) {
      const diagramXml = await zip.file(diagramPath)?.async('string');
      if (!diagramXml) {
        log('warn', `Failed to read diagram: ${diagramPath}`);
        continue;
      }

      const diagramId = diagramPath.match(/data(\d+)\.xml/)?.[1] || '1';
      const relsMap = await parseDiagramRels(zip, diagramId);

      const diagram = parseDiagram(diagramXml, diagramId, relsMap);
      diagrams.push(diagram);
    }

    return diagrams;
  } catch (error) {
    log('error', 'Failed to parse diagrams', error);
    return [];
  }
}

/**
 * 解析单个 Diagram
 * @param diagramXml Diagram XML
 * @param diagramId Diagram ID
 * @param relsMap 关联关系映射表
 * @returns Diagram 解析结果
 */
function parseDiagram(diagramXml: string, diagramId: string, relsMap: RelsMap): DiagramResult {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(diagramXml, 'application/xml');
    const root = doc.documentElement;

    // 解析类型
    const diagramType = root.getAttribute('type') || 'unknown';

    // 解析布局
    const layout = root.getAttribute('layout') || '';

    // 解析形状
    const shapes = parseDiagramShapes(root);

    log('info', `Parsed diagram ${diagramId}: type=${diagramType}, shapes=${shapes.length}`);

    return {
      id: `diagram_${diagramId}`,
      diagramType,
      layout,
      shapes,
      relsMap
    };
  } catch (error) {
    log('error', `Failed to parse diagram ${diagramId}`, error);
    return {
      id: `diagram_${diagramId}`,
      relsMap
    };
  }
}

/**
 * 解析 Diagram 形状
 */
function parseDiagramShapes(root: Element): DiagramShape[] {
  const shapes: DiagramShape[] = [];

  const spTree = getFirstChildByTagNS(root, 'spTree', NS.d);
  if (!spTree) return shapes;

  const spNodes = Array.from(spTree.children).filter(
    child => child.tagName.includes(':sp') || child.localName === 'sp'
  );

  for (const sp of spNodes) {
    const id = getAttrSafe(sp, 'id', `shape_${Date.now()}`);
    const type = sp.getAttribute('type') || 'shape';

    const xfrm = getFirstChildByTagNS(sp, 'xfrm', NS.a);
    let position, size;

    if (xfrm) {
      const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
      const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

      if (off) {
        const x = parseInt(off.getAttribute('x') || '0');
        const y = parseInt(off.getAttribute('y') || '0');
        position = { x: emu2px(x), y: emu2px(y) };
      }

      if (ext) {
        const width = parseInt(ext.getAttribute('cx') || '0');
        const height = parseInt(ext.getAttribute('cy') || '0');
        size = { width: emu2px(width), height: emu2px(height) };
      }
    }

    // 解析文本
    const text = parseShapeText(sp);

    shapes.push({
      id,
      type,
      position,
      size,
      text
    });
  }

  return shapes;
}

/**
 * 解析形状文本
 */
function parseShapeText(sp: Element): string | undefined {
  const txBody = getFirstChildByTagNS(sp, 'txBody', NS.d);
  if (!txBody) return undefined;

  const p = getFirstChildByTagNS(txBody, 'p', NS.a);
  if (!p) return undefined;

  const text = extractTextFromParagraph(p);
  return text;
}

/**
 * 解析图表关联关系
 */
async function parseChartRels(zip: JSZip, chartId: string): Promise<RelsMap> {
  try {
    const relsPath = `ppt/charts/_rels/${chartId}.xml.rels`;
    const relsXml = await zip.file(relsPath)?.async('string');

    if (!relsXml) {
      log('warn', `Chart rels not found: ${relsPath}`);
      return {};
    }

    return parseRelationshipsXml(relsXml);
  } catch (error) {
    log('warn', `Failed to parse chart rels for ${chartId}`, error);
    return {};
  }
}

/**
 * 解析 Diagram 关联关系
 */
async function parseDiagramRels(zip: JSZip, diagramId: string): Promise<RelsMap> {
  try {
    const relsPath = `ppt/diagrams/_rels/data${diagramId}.xml.rels`;
    const relsXml = await zip.file(relsPath)?.async('string');

    if (!relsXml) {
      log('warn', `Diagram rels not found: ${relsPath}`);
      return {};
    }

    return parseRelationshipsXml(relsXml);
  } catch (error) {
    log('warn', `Failed to parse diagram rels for ${diagramId}`, error);
    return {};
  }
}

/**
 * 解析关联关系XML
 */
function parseRelationshipsXml(relsXml: string): RelsMap {
  const relsMap: RelsMap = {};

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(relsXml, 'application/xml');
    const root = doc.documentElement;

    const relationships = Array.from(root.children).filter(
      child => child.tagName === 'Relationship' || child.tagName.includes(':Relationship')
    );

    for (const rel of relationships) {
      const id = rel.getAttribute('Id') || '';
      const type = rel.getAttribute('Type') || '';
      const target = rel.getAttribute('Target') || '';

      if (id) {
        relsMap[id.replace('rId', '')] = {
          id,
          type,
          target
        };
      }
    }
  } catch (error) {
    log('warn', 'Failed to parse relationships XML', error);
  }

  return relsMap;
}
